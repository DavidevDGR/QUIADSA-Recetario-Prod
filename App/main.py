# -*- coding: utf-8 -*-
"""
Aplicación táctil de control de fabricación en cuba, con soporte de
MÚLTIPLES Órdenes de Fabricación simultáneas mediante un sistema de
pestañas.

FASE 1: Identificación de operarios + lectura OF + cuba + peso tara.
FASE 2: Ejecución de secciones con temporizador; gestión de operarios
        habilitada; Cancelar/Anterior/Siguiente.
FASE 3: Pantalla final -> peso total -> comprobación de desviación del
        peso neto respecto a la cantidad prevista en la OF -> exporta a
        Excel -> cierra la pestaña (y abre una nueva si era la única).

Ejecutar con:  python main.py
"""

import tkinter as tk
from tkinter import messagebox
import datetime
from dataclasses import dataclass, field
from typing import Optional, List, Set

import config
import export
from db import SolmicroRepository, LocalRepository, OrdenFabricacion
from onscreen_keyboard import attach_keyboard

FONT_N = (config.FONT_FAMILY, config.FONT_SIZE_NORMAL)
FONT_T = (config.FONT_FAMILY, config.FONT_SIZE_TITLE, "bold")
FONT_TIMER = (config.FONT_FAMILY, config.FONT_SIZE_TIMER, "bold")
FONT_TAB = (config.FONT_FAMILY, 12, "bold")

COLOR_TAB_ACTIVA = "#FFFFFF"
COLOR_TAB_INACTIVA = "#CFD8DC"
COLOR_TABBAR_FONDO = "#90A4AE"


@dataclass
class TabState:
    """Estado completo e independiente de una Orden de Fabricación en
    curso dentro de una pestaña. Cada pestaña tiene su propia instancia,
    lo que permite tener varias OF abiertas a la vez sin interferencias."""
    tab_id: int
    fase: str = "fase1"           # fase1 / fase2 / fase3
    operarios: List[str] = field(default_factory=list)
    orden: Optional[OrdenFabricacion] = None
    cuba_id: Optional[str] = None
    peso_tara: Optional[float] = None
    ejecucion_id: Optional[int] = None
    seccion_idx: int = 0
    seccion_actual_id: Optional[int] = None
    # Índice de sección para el que ya se ha llamado a iniciar_seccion() y
    # fijado seccion_inicio_dt. Permite distinguir "estoy mostrando de
    # nuevo la misma sección" (p.ej. al cambiar de pestaña) de "empiezo una
    # sección nueva", para no reiniciar el temporizador ni duplicar el
    # registro en cada redibujado de la pantalla.
    seccion_iniciada_idx: Optional[int] = None
    seccion_inicio_dt: Optional[datetime.datetime] = None
    hora_inicio_of: Optional[datetime.datetime] = None
    # Operarios que han estado presentes en algún momento durante la
    # sección actualmente en curso (se guarda al cerrar la sección, para
    # trazabilidad/exportación, aunque luego se hayan quitado de la lista).
    seccion_operarios_historial: Set[str] = field(default_factory=set)
    # id de la pausa actualmente abierta (None si no hay ninguna en curso)
    pausa_actual_id: Optional[int] = None

    def etiqueta(self) -> str:
        if self.orden:
            return f"{self.orden.codigo_of}"
        return "Nueva OF"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Control de Fabricación - Cuba")
        self.attributes("-fullscreen", True)  # Modo táctil pantalla completa
        self.bind("<Escape>", lambda e: self.attributes("-fullscreen", False))

        self.solmicro = SolmicroRepository()
        self.local_db = LocalRepository()

        self.timer_job = None

        # --- Estado de pestañas (una TabState por Orden de Fabricación) ---
        self.tabs = {}          # tab_id -> TabState
        self.tab_order = []     # orden visual de las pestañas
        self.active_tab_id = None
        self._tab_counter = 0

        # --- Barra de pestañas, estilo navegador ---
        self.tabbar = tk.Frame(self, bg=COLOR_TABBAR_FONDO)
        self.tabbar.pack(fill="x", side="top")

        # --- Contenedor de la pantalla activa (Fase 1/2/3) ---
        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        # Si se cierra la ventana (botón X del sistema operativo) con
        # alguna OF en curso, se cancela y se exporta como CANCELADA antes
        # de cerrar de verdad la aplicación.
        self.protocol("WM_DELETE_WINDOW", self._on_cerrar_ventana)

        self._nueva_pestaña()

    # =====================================================================
    # GESTIÓN DE PESTAÑAS
    # =====================================================================
    @property
    def tab_activo(self) -> TabState:
        return self.tabs[self.active_tab_id]

    def _nueva_pestaña(self, activar=True):
        tab_id = self._tab_counter
        self._tab_counter += 1
        self.tabs[tab_id] = TabState(tab_id=tab_id)
        self.tab_order.append(tab_id)
        if activar:
            self._detener_timer()
            self.active_tab_id = tab_id
        self._refrescar_barra_pestañas()
        if activar:
            self._mostrar_pantalla_activa()
        return tab_id

    def _cambiar_pestaña(self, tab_id):
        if tab_id == self.active_tab_id:
            return
        self._detener_timer()
        self.active_tab_id = tab_id
        self._refrescar_barra_pestañas()
        self._mostrar_pantalla_activa()

    def _cerrar_pestaña(self, tab_id):
        """Cierra una pestaña (botón ✕). Si la OF está en curso, se pide
        confirmación y un motivo de cancelación antes de cerrarla."""
        t = self.tabs[tab_id]
        if t.fase in ("fase2", "fase3"):
            if not messagebox.askyesno(
                "Cerrar pestaña",
                f"La orden {t.etiqueta()} está en curso.\n"
                "Se CANCELARÁ y se registrará en el Excel. ¿Continuar?"):
                return
            if tab_id == self.active_tab_id:
                self._detener_timer()
            motivo = self._pedir_motivo_incidencia(
                f"Se va a cancelar la orden {t.etiqueta()} al cerrar su pestaña.\n"
                "Indique el motivo de la cancelación:",
                "Motivo de la cancelación (cierre de pestaña)")
            self._cancelar_of_tab(t, motivo)

        self._quitar_tab_de_listas(tab_id)

    def _quitar_tab_de_listas(self, tab_id):
        """Elimina la pestaña de las estructuras internas y refresca la
        pantalla si era la que estaba activa. Siempre deja al menos una
        pestaña disponible (se abre una nueva en blanco si se cierran
        todas, igual que mantener una ventana abierta en un navegador)."""
        idx = self.tab_order.index(tab_id)
        era_activa = (tab_id == self.active_tab_id)
        self.tab_order.remove(tab_id)
        del self.tabs[tab_id]

        if not self.tab_order:
            self._nueva_pestaña(activar=True)
            return

        if era_activa:
            nuevo_idx = min(idx, len(self.tab_order) - 1)
            self._detener_timer()
            self.active_tab_id = self.tab_order[nuevo_idx]
            self._refrescar_barra_pestañas()
            self._mostrar_pantalla_activa()
        else:
            self._refrescar_barra_pestañas()

    def _reiniciar_tab_a_fase1(self, tab_id):
        """Tras Finalizar o Cancelar una OF, la pestaña vuelve a la Fase 1
        en blanco (lista para empezar otra OF) en lugar de cerrarse. La
        pestaña se mantiene abierta con el mismo tab_id."""
        self.tabs[tab_id] = TabState(tab_id=tab_id)
        if tab_id == self.active_tab_id:
            self._detener_timer()
            self._refrescar_barra_pestañas()
            self._mostrar_pantalla_activa()
        else:
            self._refrescar_barra_pestañas()

    def _refrescar_barra_pestañas(self):
        for w in self.tabbar.winfo_children():
            w.destroy()

        for tab_id in self.tab_order:
            t = self.tabs[tab_id]
            activa = tab_id == self.active_tab_id
            celda = tk.Frame(self.tabbar, bg=COLOR_TABBAR_FONDO)
            celda.pack(side="left", padx=(4, 0), pady=4)

            bg = COLOR_TAB_ACTIVA if activa else COLOR_TAB_INACTIVA
            relieve = "sunken" if activa else "raised"
            btn = tk.Button(
                celda, text=t.etiqueta(), font=FONT_TAB, bg=bg, relief=relieve,
                command=lambda tid=tab_id: self._cambiar_pestaña(tid))
            btn.pack(side="left", ipadx=10, ipady=6)

            cerrar = tk.Button(
                celda, text="✕", font=FONT_TAB, bg=bg, fg="#B71C1C", relief=relieve,
                command=lambda tid=tab_id: self._cerrar_pestaña(tid))
            cerrar.pack(side="left", ipady=6)

        tk.Button(self.tabbar, text="+ Nueva OF", font=FONT_TAB, bg="#78909C",
                  fg="white", command=lambda: self._nueva_pestaña(activar=True)
                  ).pack(side="left", padx=6, pady=4, ipadx=8, ipady=4)

    def _mostrar_pantalla_activa(self):
        t = self.tab_activo
        if t.fase == "fase1":
            self.mostrar_fase1()
        elif t.fase == "fase2":
            self.mostrar_fase2()
        elif t.fase == "fase3":
            self.mostrar_fase3()

    def _limpiar_container(self):
        for w in self.container.winfo_children():
            w.destroy()

    # =====================================================================
    # GESTIÓN DE OPERARIOS (reutilizable en Fase 1 Y Fase 2: se puede
    # añadir/quitar operarios en cualquier momento de la ejecución)
    # =====================================================================
    def _construir_gestor_operarios(self, parent, t: TabState):
        tk.Label(parent, text="Identificador de operario (Enter o botón Añadir):",
                 font=FONT_N).pack(pady=(10, 0))
        op_frame = tk.Frame(parent)
        op_frame.pack(pady=5)
        self.entry_operario = tk.Entry(op_frame, font=FONT_N, width=20)
        self.entry_operario.pack(side="left", padx=5)
        self.entry_operario.bind("<Return>", lambda e: self._añadir_operario())
        attach_keyboard(self.entry_operario)
        tk.Button(op_frame, text="Añadir", font=FONT_N,
                  command=self._añadir_operario).pack(side="left")

        lista_frame = tk.Frame(parent)
        lista_frame.pack(pady=5)
        self.lista_operarios = tk.Listbox(lista_frame, font=FONT_N, width=25,
                                           height=4, exportselection=False)
        self.lista_operarios.pack(side="left", padx=(0, 5))
        for op in t.operarios:
            self.lista_operarios.insert(tk.END, op)
        tk.Button(lista_frame, text="Eliminar\nseleccionado", font=FONT_N,
                  bg="#e53935", fg="white",
                  command=self._eliminar_operario).pack(side="left")

    def _añadir_operario(self):
        t = self.tab_activo
        val = self.entry_operario.get().strip()
        if val and val not in t.operarios:
            t.operarios.append(val)
            self.lista_operarios.insert(tk.END, val)
            if t.fase == "fase2":
                # Se registra como "ha estado" en la sección actual, aunque
                # más tarde se le quite de la lista de operarios activos.
                t.seccion_operarios_historial.add(val)
        self.entry_operario.delete(0, tk.END)
        self.entry_operario.focus_set()

    def _eliminar_operario(self):
        t = self.tab_activo
        seleccion = self.lista_operarios.curselection()
        if not seleccion:
            messagebox.showinfo("Eliminar operario",
                                 "Seleccione un operario de la lista para eliminarlo.")
            return
        idx = seleccion[0]
        self.lista_operarios.delete(idx)
        del t.operarios[idx]

    # =====================================================================
    # FASE 1: IDENTIFICACIÓN + OF + CUBA + TARA
    # =====================================================================
    def mostrar_fase1(self):
        t = self.tab_activo
        t.fase = "fase1"
        self._limpiar_container()
        f = self.container

        tk.Label(f, text="INICIO DE ORDEN DE FABRICACIÓN", font=FONT_T).pack(pady=15)

        self._construir_gestor_operarios(f, t)

        tk.Label(f, text="Escanear Orden de Fabricación (OF):",
                 font=FONT_N).pack(pady=(15, 0))
        self.entry_of = tk.Entry(f, font=FONT_N, width=30)
        self.entry_of.pack(pady=5)
        self.entry_of.bind("<Return>", lambda e: self._leer_of())
        attach_keyboard(self.entry_of)
        self.of_info_lbl = tk.Label(f, text="", font=FONT_N, fg="green",
                                     justify="left")
        self.of_info_lbl.pack(pady=5)
        if t.orden:
            self._mostrar_info_of(t.orden)

        tk.Label(f, text="Identificador de cuba:", font=FONT_N).pack(pady=(15, 0))
        self.entry_cuba = tk.Entry(f, font=FONT_N, width=20)
        self.entry_cuba.pack(pady=5)
        if t.cuba_id:
            self.entry_cuba.insert(0, t.cuba_id)
        attach_keyboard(self.entry_cuba)

        tk.Label(f, text="Peso tara de la cuba (kg):", font=FONT_N).pack(pady=(10, 0))
        self.entry_tara = tk.Entry(f, font=FONT_N, width=20)
        self.entry_tara.pack(pady=5)
        if t.peso_tara is not None:
            self.entry_tara.insert(0, str(t.peso_tara))
        attach_keyboard(self.entry_tara, numeric_only=True)

        tk.Button(f, text="AVANZAR ▶", font=FONT_T, bg="#4CAF50", fg="white",
                  height=config.BTN_HEIGHT, command=self._validar_fase1
                  ).pack(pady=25)

        self.entry_operario.focus_set()

    def _mostrar_info_of(self, orden):
        self.of_info_lbl.config(
            text=(f"OF: {orden.codigo_of}  |  Artículo: {orden.articulo}\n"
                  f"Cantidad: {orden.cantidad_kg} kg  |  Centro: {orden.centro}\n"
                  f"Secciones: {len(orden.secciones)}")
        )

    def _leer_of(self):
        t = self.tab_activo
        codigo = self.entry_of.get().strip()
        self.entry_of.delete(0, tk.END)
        if not codigo:
            return
        orden = self.solmicro.obtener_orden_fabricacion(codigo)
        if not orden:
            messagebox.showerror("OF no encontrada",
                                  f"No se encontró la orden de fabricación '{codigo}'.")
            return
        t.orden = orden
        self._mostrar_info_of(orden)
        self._refrescar_barra_pestañas()  # la pestaña pasa a mostrar el código de OF

    def _validar_fase1(self):
        t = self.tab_activo
        if not t.operarios:
            messagebox.showwarning("Falta información",
                                    "Debe identificarse al menos un operario.")
            return
        if not t.orden:
            messagebox.showwarning("Falta información",
                                    "Debe escanear una Orden de Fabricación válida.")
            return
        cuba_id = self.entry_cuba.get().strip()
        if not cuba_id:
            messagebox.showwarning("Falta información",
                                    "Debe indicar el identificador de la cuba.")
            return
        try:
            peso_tara = float(self.entry_tara.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Peso inválido",
                                    "Introduzca el peso de tara en formato numérico.")
            return

        t.cuba_id = cuba_id
        t.peso_tara = peso_tara
        t.hora_inicio_of = datetime.datetime.now()
        t.ejecucion_id = self.local_db.crear_ejecucion(
            codigo_of=t.orden.codigo_of,
            maquina_id=t.orden.centro,
            operarios=t.operarios,
            cuba_id=t.cuba_id,
            peso_tara=t.peso_tara,
            cantidad_of_kg=t.orden.cantidad_kg,
        )
        t.seccion_idx = 0
        t.fase = "fase2"
        self.mostrar_fase2()

    # =====================================================================
    # FASE 2: EJECUCIÓN DE SECCIONES
    # =====================================================================
    def mostrar_fase2(self):
        t = self.tab_activo
        t.fase = "fase2"
        self._limpiar_container()
        f = self.container

        secciones = t.orden.secciones
        seccion = secciones[t.seccion_idx]
        es_ultima = t.seccion_idx == len(secciones) - 1

        # --- Barra superior: Cancelar + Pausa + progreso ---
        top = tk.Frame(f)
        top.pack(fill="x", pady=10, padx=10)
        tk.Button(top, text="✕ CANCELAR OF", font=FONT_N, bg="#e53935", fg="white",
                  command=self._cancelar_of).pack(side="left")

        self.btn_pausa = tk.Button(top, font=FONT_N, command=self._toggle_pausa)
        self.btn_pausa.pack(side="left", padx=(10, 0))
        self._actualizar_boton_pausa(t)

        tk.Label(top, text=f"Sección {seccion.numero} de {len(secciones)}",
                 font=FONT_N).pack(side="right")

        tk.Label(f, text=f"OF: {t.orden.codigo_of}  |  Cuba: {t.cuba_id}",
                 font=FONT_N, fg="gray").pack(pady=(0, 5))

        # --- Instrucción ---
        tk.Label(f, text=seccion.texto, font=FONT_T, wraplength=800,
                 justify="center").pack(pady=10, padx=20)

        # --- Temporizador ---
        self.timer_lbl = tk.Label(f, text="00:00", font=FONT_TIMER, fg="#1565C0")
        self.timer_lbl.pack(pady=10)
        tk.Label(f, text=f"Tiempo previsto: {seccion.tiempo_ejecucion_min} min",
                 font=FONT_N, fg="gray").pack()

        # --- Gestión de operarios, habilitada también durante la ejecución ---
        op_box = tk.LabelFrame(f, text="Operarios en esta sección", font=FONT_N)
        op_box.pack(pady=10, padx=20, fill="x")
        self._construir_gestor_operarios(op_box, t)

        # --- Navegación inferior ---
        nav = tk.Frame(f)
        nav.pack(side="bottom", fill="x", pady=20, padx=20)

        btn_anterior = tk.Button(
            nav, text="◀ ANTERIOR", font=FONT_T, height=config.BTN_HEIGHT,
            state=("disabled" if t.seccion_idx == 0 else "normal"),
            command=self._seccion_anterior
        )
        btn_anterior.pack(side="left", expand=True, fill="x", padx=10)

        texto_siguiente = "FINALIZAR ✔" if es_ultima else "SIGUIENTE ▶"
        color_siguiente = "#2E7D32" if es_ultima else "#1565C0"
        tk.Button(nav, text=texto_siguiente, font=FONT_T, bg=color_siguiente,
                  fg="white", height=config.BTN_HEIGHT,
                  command=self._seccion_siguiente
                  ).pack(side="right", expand=True, fill="x", padx=10)

        # Solo se arranca el registro de la sección (y su temporizador) la
        # PRIMERA vez que se muestra esta sección. Si la pantalla se
        # redibuja por otro motivo (p.ej. cambiar de pestaña y volver, o
        # cambiar a otra pestaña y regresar a esta misma sección) NO se
        # reinicia nada, así el tiempo transcurrido se mantiene correcto.
        if t.seccion_iniciada_idx != t.seccion_idx:
            t.seccion_operarios_historial = set(t.operarios)
            t.seccion_actual_id = self.local_db.iniciar_seccion(
                ejecucion_of_id=t.ejecucion_id,
                numero_seccion=seccion.numero,
                texto=seccion.texto,
                tiempo_previsto_min=seccion.tiempo_ejecucion_min,
            )
            t.seccion_inicio_dt = datetime.datetime.now()
            t.seccion_iniciada_idx = t.seccion_idx

        self._tick_timer()

    def _tick_timer(self):
        # Solo debe haber un temporizador visual activo a la vez: el de la
        # pestaña actualmente mostrada en pantalla. Las demás pestañas
        # siguen "corriendo" igualmente, porque su tiempo se calcula a
        # partir de seccion_inicio_dt (guardado en su TabState) en cuanto
        # se vuelve a mostrar esa pestaña, sin necesidad de un temporizador
        # en segundo plano por cada una.
        t = self.tab_activo
        if t.fase != "fase2" or t.seccion_inicio_dt is None:
            return
        transcurrido = (datetime.datetime.now() - t.seccion_inicio_dt).total_seconds()
        mins, secs = divmod(int(transcurrido), 60)
        self.timer_lbl.config(text=f"{mins:02d}:{secs:02d}")
        self.timer_job = self.after(1000, self._tick_timer)

    def _detener_timer(self):
        if self.timer_job:
            self.after_cancel(self.timer_job)
            self.timer_job = None

    def _tiempo_transcurrido_min(self, t: TabState):
        return (datetime.datetime.now() - t.seccion_inicio_dt).total_seconds() / 60.0

    def _fuera_de_margen_tiempo(self, t: TabState, tiempo_previsto_min):
        """True si se avanza fuera del margen (config.MARGEN_TIEMPO_PORCENTAJE)
        antes o después de que el tiempo previsto se cumpla."""
        if tiempo_previsto_min <= 0:
            return False
        transcurrido = self._tiempo_transcurrido_min(t)
        margen = tiempo_previsto_min * config.MARGEN_TIEMPO_PORCENTAJE
        limite_inferior = tiempo_previsto_min - margen
        limite_superior = tiempo_previsto_min + margen
        return transcurrido < limite_inferior or transcurrido > limite_superior

    def _pedir_motivo_incidencia(self, mensaje: str,
                                  titulo="Motivo de la incidencia") -> str:
        """Ventana modal genérica para pedir un motivo de incidencia
        (se reutiliza para desvíos de tiempo, de peso y para motivos de
        cancelación). Es de obligada cumplimentación: no se puede cerrar
        con el botón de la ventana del sistema operativo, solo escribiendo
        el motivo y pulsando Aceptar."""
        resultado = {"motivo": None}

        win = tk.Toplevel(self)
        win.title(titulo)
        win.grab_set()
        win.attributes("-topmost", True)
        win.protocol("WM_DELETE_WINDOW", lambda: None)  # no se puede cerrar sin rellenar

        tk.Label(win, text=mensaje, font=FONT_N, wraplength=500,
                 justify="left").pack(padx=20, pady=15)

        txt = tk.Text(win, font=FONT_N, width=45, height=4)
        txt.pack(padx=20, pady=10)
        txt.focus_set()
        attach_keyboard(txt)

        def aceptar():
            motivo = txt.get("1.0", "end").strip()
            if not motivo:
                messagebox.showwarning("Motivo requerido",
                                        "Debe indicar un motivo para continuar.",
                                        parent=win)
                return
            resultado["motivo"] = motivo
            win.destroy()

        tk.Button(win, text="Aceptar", font=FONT_N, bg="#1565C0", fg="white",
                  command=aceptar).pack(pady=15)

        win.wait_window()
        return resultado["motivo"]

    def _seccion_siguiente(self):
        t = self.tab_activo

        if t.pausa_actual_id is not None:
            messagebox.showwarning(
                "Pausa sin reanudar",
                "Hay una pausa en curso que no se ha reanudado.\n"
                "Pulse REANUDAR antes de continuar a la siguiente sección.")
            return

        secciones = t.orden.secciones
        seccion = secciones[t.seccion_idx]
        es_ultima = t.seccion_idx == len(secciones) - 1

        motivo = None
        if self._fuera_de_margen_tiempo(t, seccion.tiempo_ejecucion_min):
            pct = int(config.MARGEN_TIEMPO_PORCENTAJE * 100)
            motivo = self._pedir_motivo_incidencia(
                f"El tiempo empleado se desvía más de un {pct}% del previsto.\n"
                "Indique el motivo para poder continuar:",
                "Motivo del desvío de tiempo")
            if motivo is None:
                return  # el operario canceló el diálogo, no avanza

        self._detener_timer()
        self.local_db.cerrar_seccion(t.seccion_actual_id, incidencia_motivo=motivo)
        self.local_db.registrar_operarios_seccion(
            t.seccion_actual_id, sorted(t.seccion_operarios_historial))

        if es_ultima:
            t.fase = "fase3"
            self.mostrar_fase3()
        else:
            t.seccion_idx += 1
            self.mostrar_fase2()

    def _seccion_anterior(self):
        t = self.tab_activo
        if t.seccion_idx == 0:
            return
        if t.pausa_actual_id is not None:
            messagebox.showwarning(
                "Pausa sin reanudar",
                "Hay una pausa en curso que no se ha reanudado.\n"
                "Pulse REANUDAR antes de volver a la sección anterior.")
            return
        self._detener_timer()
        self.local_db.cerrar_seccion(t.seccion_actual_id,
                                      incidencia_motivo="Vuelta a sección anterior")
        self.local_db.registrar_operarios_seccion(
            t.seccion_actual_id, sorted(t.seccion_operarios_historial))
        t.seccion_idx -= 1
        self.mostrar_fase2()

    def _actualizar_boton_pausa(self, t: TabState):
        """Refleja en el botón si hay una pausa en curso o no. Este botón
        NO detiene el temporizador de la sección; solo registra la fecha/
        hora de inicio y fin de la pausa (y su duración) para el Excel."""
        if t.pausa_actual_id is None:
            self.btn_pausa.config(text="⏸ PAUSA", bg="#FBC02D", fg="black")
        else:
            self.btn_pausa.config(text="▶ REANUDAR", bg="#43A047", fg="white")

    def _toggle_pausa(self):
        t = self.tab_activo
        if t.pausa_actual_id is None:
            t.pausa_actual_id = self.local_db.iniciar_pausa(
                t.ejecucion_id, ejecucion_seccion_id=t.seccion_actual_id)
        else:
            self.local_db.finalizar_pausa(t.pausa_actual_id)
            t.pausa_actual_id = None
        self._actualizar_boton_pausa(t)

    def _finalizar_pausa_si_procede(self, t: TabState):
        """Si se cancela/finaliza la OF con una pausa todavía abierta, se
        cierra automáticamente para no dejar datos incompletos en el Excel."""
        if t.pausa_actual_id is not None:
            self.local_db.finalizar_pausa(t.pausa_actual_id)
            t.pausa_actual_id = None

    def _cancelar_of_tab(self, t: "TabState", motivo: str):
        """Cancela la ejecución de la OF de la pestaña indicada. Si la
        sección aún estaba en curso (Fase 2, sin cerrar todavía), se cierra
        ahora registrando el motivo de la cancelación en ESA fila concreta
        para que quede reflejado en el Excel."""
        self._finalizar_pausa_si_procede(t)
        if t.fase == "fase2" and t.seccion_actual_id:
            self.local_db.cerrar_seccion(t.seccion_actual_id, incidencia_motivo=motivo)
            self.local_db.registrar_operarios_seccion(
                t.seccion_actual_id, sorted(t.seccion_operarios_historial))
        self.local_db.cancelar_ejecucion(t.ejecucion_id)
        self._exportar_tab(t)

    def _on_cerrar_ventana(self):
        """Se llama al pulsar el botón de cerrar de la ventana (SO). Si hay
        una o varias OF en curso (en cualquier pestaña), se piden
        confirmación y motivo, se cancelan y se exportan como CANCELADA
        antes de cerrar de verdad la aplicación."""
        en_curso = [t for t in self.tabs.values() if t.fase in ("fase2", "fase3")]

        if en_curso:
            etiquetas = ", ".join(t.etiqueta() for t in en_curso)
            if not messagebox.askyesno(
                "Cerrar aplicación",
                f"Hay {len(en_curso)} orden(es) en curso: {etiquetas}.\n"
                "Se CANCELARÁN y se registrarán en el Excel. ¿Continuar?"):
                return
            self._detener_timer()
            for t in en_curso:
                motivo = self._pedir_motivo_incidencia(
                    f"Se va a cancelar la orden {t.etiqueta()} porque se está "
                    "cerrando la aplicación.\nIndique el motivo de la cancelación:",
                    "Motivo de la cancelación (cierre de la aplicación)")
                self._cancelar_of_tab(t, motivo)

        self.destroy()

    def _cancelar_of(self):
        t = self.tab_activo
        if not messagebox.askyesno("Cancelar OF",
                                    "¿Seguro que desea cancelar la orden de fabricación?\n"
                                    "Se registrará como CANCELADA."):
            return
        self._detener_timer()
        motivo = self._pedir_motivo_incidencia(
            "Indique el motivo de la cancelación de esta orden de fabricación:",
            "Motivo de la cancelación")
        self._cancelar_of_tab(t, motivo)
        self._reiniciar_tab_a_fase1(t.tab_id)

    # =====================================================================
    # FASE 3: PESO FINAL + EXPORTACIÓN
    # =====================================================================
    def mostrar_fase3(self):
        t = self.tab_activo
        t.fase = "fase3"
        self._limpiar_container()
        f = self.container

        tiempo_total = (datetime.datetime.now() - t.hora_inicio_of).total_seconds() / 60.0

        tk.Label(f, text="ORDEN FINALIZADA", font=FONT_T, fg="#2E7D32").pack(pady=30)
        tk.Label(f, text=f"OF: {t.orden.codigo_of}", font=FONT_N).pack(pady=5)
        tk.Label(f, text=f"Tiempo total: {tiempo_total:.1f} minutos",
                 font=FONT_N).pack(pady=5)
        tk.Label(f, text=f"Peso tara: {t.peso_tara} kg", font=FONT_N).pack(pady=5)
        tk.Label(f, text=f"Cantidad prevista en la OF: {t.orden.cantidad_kg} kg",
                 font=FONT_N, fg="gray").pack(pady=2)

        tk.Label(f, text="Peso total (cuba llena) en kg:", font=FONT_N).pack(pady=(25, 0))
        entry_peso_llena = tk.Entry(f, font=FONT_N, width=20)
        entry_peso_llena.pack(pady=10)
        attach_keyboard(entry_peso_llena, numeric_only=True)
        entry_peso_llena.focus_set()

        peso_neto_lbl = tk.Label(f, text="Peso neto: -", font=FONT_N, fg="blue")
        peso_neto_lbl.pack(pady=5)

        def _actualizar_neto(event=None):
            try:
                peso_llena = float(entry_peso_llena.get().replace(",", "."))
                neto = peso_llena - t.peso_tara
                peso_neto_lbl.config(text=f"Peso neto: {neto:.2f} kg")
            except ValueError:
                peso_neto_lbl.config(text="Peso neto: -")

        entry_peso_llena.bind("<KeyRelease>", _actualizar_neto)
        entry_peso_llena.bind("<<VirtualKeyboardInput>>", _actualizar_neto)

        def finalizar():
            try:
                peso_llena = float(entry_peso_llena.get().replace(",", "."))
            except ValueError:
                messagebox.showwarning("Peso inválido",
                                        "Introduzca el peso en formato numérico.")
                return

            peso_neto = peso_llena - t.peso_tara
            motivo_peso = None
            objetivo = t.orden.cantidad_kg
            if objetivo and objetivo > 0:
                margen = objetivo * config.MARGEN_PESO_PORCENTAJE
                if peso_neto < objetivo - margen or peso_neto > objetivo + margen:
                    pct = int(config.MARGEN_PESO_PORCENTAJE * 100)
                    motivo_peso = self._pedir_motivo_incidencia(
                        f"El peso neto ({peso_neto:.2f} kg) se desvía más de un "
                        f"{pct}% de la cantidad prevista en la OF ({objetivo} kg).\n"
                        "Indique el motivo para poder finalizar:",
                        "Motivo de la desviación de peso")
                    if motivo_peso is None:
                        return  # el operario canceló, no finaliza

            self._finalizar_pausa_si_procede(t)
            self.local_db.finalizar_ejecucion(t.ejecucion_id, peso_llena,
                                               peso_incidencia_motivo=motivo_peso)
            self._exportar_tab(t)
            self._reiniciar_tab_a_fase1(t.tab_id)

        tk.Button(f, text="FIN", font=FONT_T, bg="#2E7D32", fg="white",
                  height=config.BTN_HEIGHT, command=finalizar).pack(pady=25)

    def _exportar_tab(self, t: TabState):
        ejecucion = self.local_db.obtener_ejecucion(t.ejecucion_id)
        secciones = self.local_db.obtener_secciones(t.ejecucion_id)
        pausas = self.local_db.obtener_pausas(t.ejecucion_id)
        try:
            ruta = export.exportar_ejecucion(ejecucion, secciones, pausas)
            print(f"[Exportación OK] {ruta} (usuario: {config.USUARIO_EXPORTACION})")
        except Exception as e:
            messagebox.showerror("Error de exportación",
                                  f"No se pudo exportar el Excel:\n{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
