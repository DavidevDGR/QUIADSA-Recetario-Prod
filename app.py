import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime
import pandas as pd
import os

##### SIMULACIÓN DE BASE DE DATOS EXISTENTE
DB_ORDENES_FABRICACION = {
    "OF-1001": {
        "id_maquina": "MAQ-CHRONOS-01",
        "instrucciones": [
            {
                "paso": 1,
                "texto": "Preparar materia prima y verificar moldes.",
                "tiempo_min": 1,
            },
            {
                "paso": 2,
                "texto": "Mezclado a alta revolución y control de temperatura.",
                "tiempo_min": 2,
            },
            {
                "paso": 3,
                "texto": "Limpieza del área y precalentamiento de extrusora.",
                "tiempo_min": 1,
            },
        ],
    },
    "OF-1002": {
        "id_maquina": "MAQ-VULCANO-02",
        "instrucciones": [
            {
                "paso": 1,
                "texto": "Corte de perfiles de aluminio según plano.",
                "tiempo_min": 5,
            },
            {"paso": 2, "texto": "Ensamblado de piezas y remachado.", "tiempo_min": 10},
        ],
    },
}
EXCEL_UNICO = "Registro_Produccion.xlsx"


class AppFabricacion(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Control de Órdenes de Fabricación")
        self.geometry("800x500")
        self.configure(bg="#2c3e50")
        self.operarios = []
        self.of_actual = None
        self.datos_of = None
        self.paso_actual = 0
        self.fecha_hora_inicio = None
        self.tiempo_restante_segundos = 0
        self.timer_corriendo = False
        self.fuera_de_tiempo = False
        self.historial_pasos = {}
        self.segundos_acumulados_paso = 0
        self._timer_id = (
            None  ##### Variable crítica: Almacena la referencia del loop del reloj
        )
        self.pantalla_inicio_unificada()

    def limpiar_pantalla(self):
        for widget in self.winfo_children():
            widget.destroy()

    def detener_loop_reloj(self):
        """Cancela de forma segura el bucle del reloj actual para que no se duplique."""
        self.timer_corriendo = False
        if self._timer_id is not None:
            self.after_cancel(self._timer_id)
            self._timer_id = None

    ##### PANTALLA UNIFICADA (LOGIN + OF)
    def pantalla_inicio_unificada(self):
        self.limpiar_pantalla()
        self.operarios = []
        self.of_actual = None
        self.datos_of = None
        lbl_titulo = tk.Label(
            self,
            text="INICIO DE JORNADA / ORDEN DE TRABAJO",
            font=("Arial", 22, "bold"),
            bg="#2c3e50",
            fg="white",
        )
        lbl_titulo.pack(pady=25)
        frame_form = tk.Frame(
            self, bg="#34495e", padx=20, pady=20, relief="groove", bd=2
        )
        frame_form.pack(pady=10, padx=40, fill="x")
        lbl_ops = tk.Label(
            frame_form,
            text="1. Identificador de Operarios (separados por comas):",
            font=("Arial", 13, "bold"),
            bg="#34495e",
            fg="#ecf0f1",
        )
        lbl_ops.grid(row=0, column=0, sticky="w", pady=5)
        entry_operarios = tk.Entry(frame_form, font=("Arial", 16), width=45)
        entry_operarios.grid(row=1, column=0, pady=(0, 15), sticky="w")
        entry_operarios.focus()
        lbl_of = tk.Label(
            frame_form,
            text="2. Introduzca la Orden de Fabricación (OF):",
            font=("Arial", 13, "bold"),
            bg="#34495e",
            fg="#ecf0f1",
        )
        lbl_of.grid(row=2, column=0, sticky="w", pady=5)
        entry_of = tk.Entry(frame_form, font=("Arial", 16), width=25)
        entry_of.grid(row=3, column=0, pady=(0, 5), sticky="w")
        btn_iniciar = tk.Button(
            self,
            text="🚀 INICIAR RECETA",
            font=("Arial", 16, "bold"),
            bg="#2ecc71",
            fg="white",
            padx=30,
            pady=12,
            command=lambda: self.procesar_inicio(
                entry_operarios.get(), entry_of.get().strip()
            ),
        )
        btn_iniciar.pack(pady=25)

    def procesar_inicio(self, txt_operarios, codigo_of):
        if not txt_operarios.strip():
            messagebox.showwarning(
                "Error", "Debes introducir al menos un identificador de operario."
            )
            return
        if not codigo_of:
            messagebox.showwarning(
                "Error", "Debes introducir el código de la Orden de Fabricación (OF)."
            )
            return
        if codigo_of in DB_ORDENES_FABRICACION:
            self.operarios = [
                op.strip() for op in txt_operarios.split(",") if op.strip()
            ]
            self.of_actual = codigo_of
            self.datos_of = DB_ORDENES_FABRICACION[codigo_of]
            self.paso_actual = 0
            self.historial_pasos = {}
            self.fecha_hora_inicio = datetime.now()
            self.pantalla_proceso()
        else:
            messagebox.showerror(
                "Error",
                f"La Orden de Fabricación '{codigo_of}' no existe en la Base de Datos.",
            )

    ##### PANTALLA DE PROCESO
    def pantalla_proceso(self):
        self.detener_loop_reloj()
        self.limpiar_pantalla()
        frame_top = tk.Frame(self, bg="#2c3e50")
        frame_top.pack(fill="x", padx=10, pady=10)
        btn_cancelar = tk.Button(
            frame_top,
            text="❌ Cancelar Receta",
            font=("Arial", 11, "bold"),
            bg="#e74c3c",
            fg="white",
            command=self.cancelar_receta,
        )
        btn_cancelar.pack(side="left")
        lbl_of_info = tk.Label(
            frame_top,
            text=f"OF: {self.of_actual} | Maq: {self.datos_of['id_maquina']}",
            font=("Arial", 12, "bold"),
            bg="#2c3e50",
            fg="#bdc3c7",
        )
        lbl_of_info.pack(side="right")
        instruccion_actual = self.datos_of["instrucciones"][self.paso_actual]
        lbl_paso = tk.Label(
            self,
            text=f"INSTRUCCIÓN {self.paso_actual + 1} de {len(self.datos_of['instrucciones'])}",
            font=("Arial", 16, "bold"),
            bg="#2c3e50",
            fg="#3498db",
        )
        lbl_paso.pack(pady=10)
        lbl_texto = tk.Label(
            self,
            text=instruccion_actual["texto"],
            font=("Arial", 18),
            bg="#34495e",
            fg="white",
            wraplength=700,
            justify="center",
            width=60,
            height=4,
            relief="groove",
            bd=2,
        )
        lbl_texto.pack(pady=15)
        self.lbl_timer = tk.Label(
            self, text="00:00", font=("Arial", 36, "bold"), bg="#2c3e50", fg="#2ecc71"
        )
        self.lbl_timer.pack(pady=10)
        frame_botones = tk.Frame(self, bg="#2c3e50")
        frame_botones.pack(side="bottom", fill="x", padx=20, pady=20)
        self.btn_anterior = tk.Button(
            frame_botones,
            text="⬅ Anterior",
            font=("Arial", 14, "bold"),
            bg="#95a5a6",
            fg="white",
            width=12,
            height=2,
            command=self.retroceder_paso,
        )
        self.btn_anterior.pack(side="left")
        if self.paso_actual == 0:
            self.btn_anterior.config(state="disabled")
        texto_sig = (
            "Finalizar 🏁"
            if self.paso_actual == len(self.datos_of["instrucciones"]) - 1
            else "Siguiente ➡"
        )
        color_sig = (
            "#27ae60"
            if self.paso_actual == len(self.datos_of["instrucciones"]) - 1
            else "#3498db"
        )
        self.btn_siguiente = tk.Button(
            frame_botones,
            text=texto_sig,
            font=("Arial", 14, "bold"),
            bg=color_sig,
            fg="white",
            width=12,
            height=2,
            command=self.avanzar_paso,
        )
        self.btn_siguiente.pack(side="right")
        if self.paso_actual not in self.historial_pasos:
            self.historial_pasos[self.paso_actual] = {
                "tiempo_empleado_seg": 0,
                "incidencia": "",
            }
            self.tiempo_restante_segundos = instruccion_actual["tiempo_min"] * 60
            self.fuera_de_tiempo = False
        else:
            self.tiempo_restante_segundos = max(
                0,
                (instruccion_actual["tiempo_min"] * 60)
                - self.historial_pasos[self.paso_actual]["tiempo_empleado_seg"],
            )
            self.fuera_de_tiempo = self.tiempo_restante_segundos == 0
        self.segundos_acumulados_paso = self.historial_pasos[self.paso_actual][
            "tiempo_empleado_seg"
        ]
        self.timer_corriendo = True
        self.actualizar_cronometro()

    def actualizar_cronometro(self):
        if not self.timer_corriendo:
            return
        self.segundos_acumulados_paso += 1
        self.historial_pasos[self.paso_actual][
            "tiempo_empleado_seg"
        ] = self.segundos_acumulados_paso
        if not self.fuera_de_tiempo:
            self.tiempo_restante_segundos -= 1
            minutos = self.tiempo_restante_segundos // 60
            segundos = self.tiempo_restante_segundos % 60
            self.lbl_timer.config(text=f"{minutos:02d}:{segundos:02d}", fg="#2ecc71")
            if self.tiempo_restante_segundos <= 0:
                self.fuera_de_tiempo = True
        else:
            segundos_exceso = self.segundos_acumulados_paso - (
                self.datos_of["instrucciones"][self.paso_actual]["tiempo_min"] * 60
            )
            minutos = segundos_exceso // 60
            segundos = segundos_exceso % 60
            self.lbl_timer.config(text=f"-{minutos:02d}:{segundos:02d}", fg="#e74c3c")
        self._timer_id = self.after(1000, self.actualizar_cronometro)

    def avanzar_paso(self):
        self.detener_loop_reloj()
        tiempo_limite_seg = (
            self.datos_of["instrucciones"][self.paso_actual]["tiempo_min"] * 60
        )
        if (
            self.segundos_acumulados_paso > tiempo_limite_seg
            and not self.historial_pasos[self.paso_actual]["incidencia"]
        ):
            incidencia = simpledialog.askstring(
                "Incidencia Obligatoria",
                "Se ha excedido el tiempo estipulado.\nPor favor, escribe el motivo de la incidencia:",
                parent=self,
            )
            if not incidencia or not incidencia.strip():
                messagebox.showwarning(
                    "Obligatorio", "Debe describir la incidencia para poder continuar."
                )
                self.timer_corriendo = True
                self.actualizar_cronometro()
                return
            self.historial_pasos[self.paso_actual]["incidencia"] = incidencia.strip()
        if self.paso_actual == len(self.datos_of["instrucciones"]) - 1:
            self.exportar_reporte(estado="FINALIZADA")
        else:
            self.paso_actual += 1
            self.pantalla_proceso()

    def retroceder_paso(self):
        self.detener_loop_reloj()
        if self.paso_actual > 0:
            self.paso_actual -= 1
            self.pantalla_proceso()

    def cancelar_receta(self):
        if messagebox.askyesno(
            "Confirmar Cancelación",
            "¿Seguro que deseas cancelar la receta?\nSe guardará una única fila de cancelación en el archivo histórico.",
        ):
            self.detener_loop_reloj()
            self.exportar_reporte(estado="CANCELADA")

    ##### LOGICA DE EXPORTACIÓN
    def exportar_reporte(self, estado="FINALIZADA"):
        self.detener_loop_reloj()
        fecha_hora_final = datetime.now()
        tiempo_total_of_seg = sum(
            item["tiempo_empleado_seg"] for item in self.historial_pasos.values()
        )
        tiempos_individuales_lista = []
        incidencias_lista = []
        for index, inst in enumerate(self.datos_of["instrucciones"]):
            datos_paso = self.historial_pasos.get(
                index, {"tiempo_empleado_seg": 0, "incidencia": "No alcanzado"}
            )
            minutos_paso = round(datos_paso["tiempo_empleado_seg"] / 60, 2)
            tiempos_individuales_lista.append(f"P{inst['paso']}: {minutos_paso} min")
            txt_incidencia = (
                datos_paso["incidencia"] if datos_paso["incidencia"] else "Ninguna"
            )
            incidencias_lista.append(f"P{inst['paso']}: {txt_incidencia}")
        tiempos_por_instruccion_str = " | ".join(tiempos_individuales_lista)
        incidencias_consolidadas_str = " | ".join(incidencias_lista)
        nueva_fila_of = {
            "Fecha Registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Estado Receta": estado,
            "OF": self.of_actual,
            "ID Máquina": self.datos_of["id_maquina"],
            "Operarios": ", ".join(self.operarios),
            "Fecha/Hora Inicio OF": self.fecha_hora_inicio.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            "Fecha/Hora Fin/Cancelación": fecha_hora_final.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            "Tiempos por Instrucción": tiempos_por_instruccion_str,
            "Tiempo Total Receta (Min)": round(tiempo_total_of_seg / 60, 2),
            "Incidencias Registradas": incidencias_consolidadas_str,
        }
        df_nueva_fila = pd.DataFrame([nueva_fila_of])
        try:
            if os.path.exists(EXCEL_UNICO):
                df_existente = pd.read_excel(EXCEL_UNICO)
                df_consolidado = pd.concat(
                    [df_existente, df_nueva_fila], ignore_index=True
                )
            else:
                df_consolidado = df_nueva_fila
            df_consolidado.to_excel(EXCEL_UNICO, index=False)
            if estado == "CANCELADA":
                messagebox.showwarning(
                    "Receta Cancelada",
                    f"Orden archivada como cancelada en una única fila dentro de:\n{EXCEL_UNICO}",
                )
            else:
                messagebox.showinfo(
                    "Éxito", f"Orden completada. Guardada nueva fila en:\n{EXCEL_UNICO}"
                )
        except Exception as e:
            messagebox.showerror(
                "Error de Almacenamiento",
                f"No se pudo escribir en el archivo único {EXCEL_UNICO}.\nMotivo: {str(e)}",
            )
        self.pantalla_inicio_unificada()


if __name__ == "__main__":
    app = AppFabricacion()
    app.mainloop()
