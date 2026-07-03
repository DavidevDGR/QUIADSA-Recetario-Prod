# -*- coding: utf-8 -*-
"""
Aplicación táctil de control de fabricación en cuba.

FASE 1: Identificación de operarios + lectura OF + cuba + peso tara.
FASE 2: Ejecución de secciones con temporizador, Cancelar/Anterior/Siguiente.
FASE 3: Pantalla final -> peso llena -> exporta a Excel -> vuelve a Fase 1.

Ejecutar con:  python main.py
"""

import tkinter as tk
from tkinter import messagebox
import datetime

import config
from db import SolmicroRepository, LocalRepository, OrdenFabricacion
from onscreen_keyboard import attach_keyboard

FONT_N = (config.FONT_FAMILY, config.FONT_SIZE_NORMAL)
FONT_T = (config.FONT_FAMILY, config.FONT_SIZE_TITLE, "bold")
FONT_TIMER = (config.FONT_FAMILY, config.FONT_SIZE_TIMER, "bold")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Control de Fabricación - Cuba")
        self.attributes("-fullscreen", True)  # Modo táctil pantalla completa
        self.bind("<Escape>", lambda e: self.attributes("-fullscreen", False))

        self.solmicro = SolmicroRepository()
        self.local_db = LocalRepository()

        # Estado de la ejecución actual
        self.reset_estado()

        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.mostrar_fase1()

    def reset_estado(self):
        self.operarios = []
        self.orden: OrdenFabricacion = None
        self.cuba_id = None
        self.peso_tara = None
        self.ejecucion_id = None
        self.seccion_idx = 0
        self.seccion_actual_id = None
        self.seccion_inicio_dt = None
        self.timer_job = None
        self.hora_inicio_of = None

    def _limpiar_container(self):
        for w in self.container.winfo_children():
            w.destroy()

    # =====================================================================
    # FASE 1: IDENTIFICACIÓN + OF + CUBA + TARA
    # =====================================================================
    def mostrar_fase1(self):
        self.reset_estado()
        self._limpiar_container()
        f = self.container

        tk.Label(f, text="INICIO DE ORDEN DE FABRICACIÓN", font=FONT_T).pack(pady=20)

        # --- Operarios ---
        tk.Label(f, text="Identificador de operario (Enter o botón Añadir):",
                 font=FONT_N).pack(pady=(10, 0))
        op_frame = tk.Frame(f)
        op_frame.pack(pady=5)
        self.entry_operario = tk.Entry(op_frame, font=FONT_N, width=20)
        self.entry_operario.pack(side="left", padx=5)
        self.entry_operario.bind("<Return>", lambda e: self._añadir_operario())
        attach_keyboard(self.entry_operario)
        tk.Button(op_frame, text="Añadir", font=FONT_N,
                  command=self._añadir_operario).pack(side="left")

        # Lista de operarios añadidos, con posibilidad de eliminarlos
        lista_frame = tk.Frame(f)
        lista_frame.pack(pady=5)
        self.lista_operarios = tk.Listbox(lista_frame, font=FONT_N, width=25,
                                           height=4, exportselection=False)
        self.lista_operarios.pack(side="left", padx=(0, 5))
        tk.Button(lista_frame, text="Eliminar\nseleccionado", font=FONT_N,
                  bg="#e53935", fg="white",
                  command=self._eliminar_operario).pack(side="left")

        # --- Orden de fabricación (código de barras -> teclado + Enter) ---
        tk.Label(f, text="Escanear Orden de Fabricación (OF):",
                 font=FONT_N).pack(pady=(20, 0))
        self.entry_of = tk.Entry(f, font=FONT_N, width=30)
        self.entry_of.pack(pady=5)
        self.entry_of.bind("<Return>", lambda e: self._leer_of())
        attach_keyboard(self.entry_of)
        self.of_info_lbl = tk.Label(f, text="", font=FONT_N, fg="green",
                                     justify="left")
        self.of_info_lbl.pack(pady=5)

        # --- Cuba ---
        tk.Label(f, text="Identificador de cuba:", font=FONT_N).pack(pady=(20, 0))
        self.entry_cuba = tk.Entry(f, font=FONT_N, width=20)
        self.entry_cuba.pack(pady=5)
        attach_keyboard(self.entry_cuba)

        # --- Peso tara ---
        tk.Label(f, text="Peso tara de la cuba (kg):", font=FONT_N).pack(pady=(10, 0))
        self.entry_tara = tk.Entry(f, font=FONT_N, width=20)
        self.entry_tara.pack(pady=5)
        attach_keyboard(self.entry_tara, numeric_only=True)

        # --- Avanzar ---
        tk.Button(f, text="AVANZAR ▶", font=FONT_T, bg="#4CAF50", fg="white",
                  height=config.BTN_HEIGHT, command=self._validar_fase1
                  ).pack(pady=30)

        self.entry_operario.focus_set()

    def _añadir_operario(self):
        val = self.entry_operario.get().strip()
        if val and val not in self.operarios:
            self.operarios.append(val)
            self.lista_operarios.insert(tk.END, val)
        self.entry_operario.delete(0, tk.END)
        self.entry_operario.focus_set()

    def _eliminar_operario(self):
        seleccion = self.lista_operarios.curselection()
        if not seleccion:
            messagebox.showinfo("Eliminar operario",
                                 "Seleccione un operario de la lista para eliminarlo.")
            return
        idx = seleccion[0]
        self.lista_operarios.delete(idx)
        del self.operarios[idx]

    def _leer_of(self):
        codigo = self.entry_of.get().strip()
        self.entry_of.delete(0, tk.END)
        if not codigo:
            return
        orden = self.solmicro.obtener_orden_fabricacion(codigo)
        if not orden:
            messagebox.showerror("OF no encontrada",
                                  f"No se encontró la orden de fabricación '{codigo}'.")
            return
        self.orden = orden
        self.of_info_lbl.config(
            text=(f"OF: {orden.codigo_of}  |  Artículo: {orden.articulo}\n"
                  f"Cantidad: {orden.cantidad_kg} kg  |  Centro: {orden.centro}\n"
                  f"Secciones: {len(orden.secciones)}")
        )

    def _validar_fase1(self):
        if not self.operarios:
            messagebox.showwarning("Falta información",
                                    "Debe identificarse al menos un operario.")
            return
        if not self.orden:
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

        self.cuba_id = cuba_id
        self.peso_tara = peso_tara

        self.hora_inicio_of = datetime.datetime.now()
        self.ejecucion_id = self.local_db.crear_ejecucion(
            codigo_of=self.orden.codigo_of,
            maquina_id=self.orden.centro,
            operarios=self.operarios,
            cuba_id=self.cuba_id,
            peso_tara=self.peso_tara,
        )
        self.seccion_idx = 0
        self.mostrar_fase2()

    # =====================================================================
    # FASE 2: EJECUCIÓN DE SECCIONES
    # =====================================================================
    def mostrar_fase2(self):
        self._limpiar_container()
        f = self.container

        secciones = self.orden.secciones
        seccion = secciones[self.seccion_idx]
        es_ultima = self.seccion_idx == len(secciones) - 1

        # --- Barra superior: Cancelar + progreso ---
        top = tk.Frame(f)
        top.pack(fill="x", pady=10, padx=10)
        tk.Button(top, text="✕ CANCELAR", font=FONT_N, bg="#e53935", fg="white",
                  command=self._cancelar_of).pack(side="left")
        tk.Label(top, text=f"Sección {seccion.numero} de {len(secciones)}",
                 font=FONT_N).pack(side="right")

        tk.Label(f, text=f"OF: {self.orden.codigo_of}  |  Cuba: {self.cuba_id}",
                 font=FONT_N, fg="gray").pack(pady=(0, 10))

        # --- Instrucción ---
        tk.Label(f, text=seccion.texto, font=FONT_T, wraplength=800,
                 justify="center").pack(pady=20, padx=20)

        # --- Temporizador ---
        self.timer_lbl = tk.Label(f, text="00:00", font=FONT_TIMER, fg="#1565C0")
        self.timer_lbl.pack(pady=20)
        tk.Label(f, text=f"Tiempo previsto: {seccion.tiempo_ejecucion_min} min",
                 font=FONT_N, fg="gray").pack()

        # --- Navegación inferior ---
        nav = tk.Frame(f)
        nav.pack(side="bottom", fill="x", pady=30, padx=20)

        btn_anterior = tk.Button(
            nav, text="◀ ANTERIOR", font=FONT_T, height=config.BTN_HEIGHT,
            state=("disabled" if self.seccion_idx == 0 else "normal"),
            command=self._seccion_anterior
        )
        btn_anterior.pack(side="left", expand=True, fill="x", padx=10)

        texto_siguiente = "FINALIZAR ✔" if es_ultima else "SIGUIENTE ▶"
        color_siguiente = "#2E7D32" if es_ultima else "#1565C0"
        tk.Button(nav, text=texto_siguiente, font=FONT_T, bg=color_siguiente,
                  fg="white", height=config.BTN_HEIGHT,
                  command=self._seccion_siguiente
                  ).pack(side="right", expand=True, fill="x", padx=10)

        # --- Arrancar registro + temporizador ---
        self.seccion_actual_id = self.local_db.iniciar_seccion(
            ejecucion_of_id=self.ejecucion_id,
            numero_seccion=seccion.numero,
            texto=seccion.texto,
            tiempo_previsto_min=seccion.tiempo_ejecucion_min,
        )
        self.seccion_inicio_dt = datetime.datetime.now()
        self._tick_timer()

    def _tick_timer(self):
        transcurrido = (datetime.datetime.now() - self.seccion_inicio_dt).total_seconds()
        mins, secs = divmod(int(transcurrido), 60)
        self.timer_lbl.config(text=f"{mins:02d}:{secs:02d}")
        self.timer_job = self.after(1000, self._tick_timer)

    def _detener_timer(self):
        if self.timer_job:
            self.after_cancel(self.timer_job)
            self.timer_job = None

    def _tiempo_transcurrido_min(self):
        return (datetime.datetime.now() - self.seccion_inicio_dt).total_seconds() / 60.0

    def _fuera_de_margen(self, tiempo_previsto_min):
        """True si se avanza un 25% antes o después del tiempo previsto."""
        if tiempo_previsto_min <= 0:
            return False
        transcurrido = self._tiempo_transcurrido_min()
        margen = tiempo_previsto_min * config.MARGEN_TIEMPO_PORCENTAJE
        limite_inferior = tiempo_previsto_min - margen
        limite_superior = tiempo_previsto_min + margen
        return transcurrido < limite_inferior or transcurrido > limite_superior

    def _pedir_motivo_incidencia(self) -> str:
        """Ventana modal simple para pedir el motivo. Devuelve el texto o None si cancela."""
        resultado = {"motivo": None}

        win = tk.Toplevel(self)
        win.title("Motivo del desvío de tiempo")
        win.grab_set()
        win.attributes("-topmost", True)

        tk.Label(win, text="El tiempo empleado se desvía más de un "
                            f"{int(config.MARGEN_TIEMPO_PORCENTAJE*100)}% del previsto.\n"
                            "Indique el motivo para poder continuar:",
                 font=FONT_N, wraplength=500, justify="left").pack(padx=20, pady=15)

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
        secciones = self.orden.secciones
        seccion = secciones[self.seccion_idx]
        es_ultima = self.seccion_idx == len(secciones) - 1

        motivo = None
        if self._fuera_de_margen(seccion.tiempo_ejecucion_min):
            motivo = self._pedir_motivo_incidencia()
            if motivo is None:
                return  # el operario canceló el diálogo, no avanza

        self._detener_timer()
        self.local_db.cerrar_seccion(self.seccion_actual_id, incidencia_motivo=motivo)

        if es_ultima:
            self.mostrar_fase3()
        else:
            self.seccion_idx += 1
            self.mostrar_fase2()

    def _seccion_anterior(self):
        if self.seccion_idx == 0:
            return
        self._detener_timer()
        # Se cierra el registro de la sección actual como "abandonada"
        self.local_db.cerrar_seccion(self.seccion_actual_id,
                                      incidencia_motivo="Vuelta a sección anterior")
        self.seccion_idx -= 1
        self.mostrar_fase2()

    def _cancelar_of(self):
        if not messagebox.askyesno("Cancelar OF",
                                    "¿Seguro que desea cancelar la orden de fabricación?\n"
                                    "Se registrará como CANCELADA."):
            return
        self._detener_timer()
        if self.seccion_actual_id:
            self.local_db.cerrar_seccion(self.seccion_actual_id,
                                          incidencia_motivo="Cancelación de OF")
        self.local_db.cancelar_ejecucion(self.ejecucion_id)
        self._exportar_y_avisar()
        self.mostrar_fase1()

    # =====================================================================
    # FASE 3: PESO FINAL + EXPORTACIÓN
    # =====================================================================
    def mostrar_fase3(self):
        self._limpiar_container()
        f = self.container

        tiempo_total = (datetime.datetime.now() - self.hora_inicio_of).total_seconds() / 60.0

        tk.Label(f, text="ORDEN FINALIZADA", font=FONT_T, fg="#2E7D32").pack(pady=30)
        tk.Label(f, text=f"OF: {self.orden.codigo_of}", font=FONT_N).pack(pady=5)
        tk.Label(f, text=f"Tiempo total: {tiempo_total:.1f} minutos",
                 font=FONT_N).pack(pady=5)
        tk.Label(f, text=f"Peso tara: {self.peso_tara} kg", font=FONT_N).pack(pady=5)

        tk.Label(f, text="Peso total (cuba llena) en kg:", font=FONT_N).pack(pady=(30, 0))
        entry_peso_llena = tk.Entry(f, font=FONT_N, width=20)
        entry_peso_llena.pack(pady=10)
        attach_keyboard(entry_peso_llena, numeric_only=True)
        entry_peso_llena.focus_set()

        peso_neto_lbl = tk.Label(f, text="Peso neto: -", font=FONT_N, fg="blue")
        peso_neto_lbl.pack(pady=5)

        def _actualizar_neto(event=None):
            try:
                peso_llena = float(entry_peso_llena.get().replace(",", "."))
                neto = peso_llena - self.peso_tara
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
            self.local_db.finalizar_ejecucion(self.ejecucion_id, peso_llena)
            self._exportar_y_avisar()
            self.mostrar_fase1()

        tk.Button(f, text="FIN", font=FONT_T, bg="#2E7D32", fg="white",
                  height=config.BTN_HEIGHT, command=finalizar).pack(pady=30)

    def _exportar_y_avisar(self):
        import export
        ejecucion = self.local_db.obtener_ejecucion(self.ejecucion_id)
        secciones = self.local_db.obtener_secciones(self.ejecucion_id)
        try:
            ruta = export.exportar_ejecucion(ejecucion, secciones)
            print(f"[Exportación OK] {ruta} (usuario: {config.USUARIO_EXPORTACION})")
        except Exception as e:
            messagebox.showerror("Error de exportación",
                                  f"No se pudo exportar el Excel:\n{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
