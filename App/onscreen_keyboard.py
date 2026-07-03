# -*- coding: utf-8 -*-
"""
Teclado virtual en pantalla (QWERTY / numérico) para uso en pantallas
táctiles sin teclado físico. Se muestra automáticamente al pulsar sobre un
Entry/Text al que se le haya aplicado `attach_keyboard()`.

NOTA: este módulo se llama "onscreen_keyboard.py" (y no "keyboard.py") a
propósito, para no chocar con el paquete de PyPI "keyboard" (atajos de
teclado globales).

Disposición basada en el diseño solicitado (grid con teclas especiales
BORRAR / ENTER / MAYÚS / ESPACIO / CERRAR en posiciones fijas), usando
tkinter.grid() con columnspan/rowspan para reproducir las celdas fusionadas.

Uso:
    entry = tk.Entry(parent, font=FONT_N)
    attach_keyboard(entry)                     # teclado completo (QWERTY)
    attach_keyboard(entry, numeric_only=True)   # solo teclado numérico
"""

import tkinter as tk

import config

# Filas de letras/dígitos del teclado QWERTY (sin las teclas especiales)
FILA_1 = list("1234567890")            # + BORRAR (columna 11)
FILA_2 = list("QWERTYUIOP")            # + ENTER (columna 11, fusionada filas 2-3)
FILA_3 = list("ASDFGHJKLÑ")
FILA_4 = list("ZXCVBNM,.-")            # + MAYUS a la izquierda (columna 0)

COLOR_TECLA = "#546E7A"
COLOR_TECLA_ACTIVA = "#78909C"
COLOR_MAYUS_ON = "#FFB300"       # resaltado bien visible cuando MAYÚS está activo
COLOR_BORRAR = "#EF5350"
COLOR_ENTER = "#1E88E5"
COLOR_CERRAR = "#66BB6A"
COLOR_ESPACIO = "#546E7A"

FONT_TECLA = (config.FONT_FAMILY, 15, "bold")


class VirtualKeyboard:
    """Teclado único y compartido por toda la aplicación (singleton)."""

    _instance = None

    def __init__(self, root: tk.Tk):
        self.root = root
        self.win = None
        self.target = None          # widget Entry/Text activo
        self.numeric_only = False
        self.mayus = False
        self.letter_buttons = {}    # char_original -> tk.Button (solo letras)
        self.mayus_button = None

    @classmethod
    def get(cls, root: tk.Tk) -> "VirtualKeyboard":
        if cls._instance is None:
            cls._instance = VirtualKeyboard(root)
        return cls._instance

    # ------------------------------------------------------------------
    def mostrar(self, target_widget, numeric_only=False):
        self.target = target_widget
        self.numeric_only = numeric_only
        self._construir_ventana()

    def ocultar(self):
        if self.win is not None:
            self.win.destroy()
            self.win = None
        self.target = None

    def _confirmar(self):
        """Botón ENTER: envía <Return> al widget activo (dispara su acción,
        p.ej. añadir operario / leer OF) y cierra el teclado."""
        if self.target is not None:
            try:
                self.target.event_generate("<Return>")
            except tk.TclError:
                pass
        self.ocultar()

    # ------------------------------------------------------------------
    def _construir_ventana(self):
        if self.win is not None:
            self.win.destroy()
        self.letter_buttons = {}
        self.mayus_button = None

        self.win = tk.Toplevel(self.root)
        self.win.overrideredirect(True)   # sin barra de título
        self.win.attributes("-topmost", True)

        ancho = self.root.winfo_screenwidth()
        alto_kb = 320 if not self.numeric_only else 260
        alto_pantalla = self.root.winfo_screenheight()
        self.win.geometry(f"{ancho}x{alto_kb}+0+{alto_pantalla - alto_kb}")
        self.win.configure(bg="#263238")

        cont = tk.Frame(self.win, bg="#263238")
        cont.pack(expand=True, fill="both", padx=6, pady=6)

        if self.numeric_only:
            self._construir_numerico(cont)
        else:
            self._construir_alfabetico(cont)

    # ------------------------------------------------------------------
    def _tecla(self, parent, texto, comando, row, col, rowspan=1, colspan=1,
               color=COLOR_TECLA, es_letra_original=None):
        btn = tk.Button(
            parent, text=texto, command=comando, font=FONT_TECLA,
            bg=color, fg="white", activebackground=COLOR_TECLA_ACTIVA,
            relief="raised", bd=2,
        )
        btn.grid(row=row, column=col, rowspan=rowspan, columnspan=colspan,
                 sticky="nsew", padx=2, pady=2, ipady=10)
        if es_letra_original is not None:
            self.letter_buttons[es_letra_original] = btn
        return btn

    # ------------------------------------------------------------------
    # QWERTY: 11 columnas x 5 filas
    #   Fila 0: 1..0 (10 cols) + BORRAR (col 10)
    #   Fila 1: Q..P (10 cols) + ENTER (col 10, rowspan 2)
    #   Fila 2: A..Ñ (10 cols)
    #   Fila 3: MAYUS(col0) + Z X C V B N M , . - (10 cols col1..10... en
    #           realidad 10 teclas => usamos col1..10, MAYUS ocupa col0
    #           EXTENDIDO: como fila0-2 tienen 10 letras/col0..9, MAYUS
    #           necesita su propia columna -> desplazamos letras a 1..10)
    #   Fila 4: CERRAR (col0..2) + ESPACIO (col3..10)
    # ------------------------------------------------------------------
    def _construir_alfabetico(self, cont):
        n_cols = 11
        n_rows = 5
        for c in range(n_cols):
            cont.grid_columnconfigure(c, weight=1, uniform="col")
        for r in range(n_rows):
            cont.grid_rowconfigure(r, weight=1)

        # Fila 0: dígitos en columnas 0-9, BORRAR en columna 10
        for i, ch in enumerate(FILA_1):
            self._tecla(cont, ch, lambda c=ch: self._pulsar(c), row=0, col=i)
        self._tecla(cont, "BORRAR", self._borrar, row=0, col=10, color=COLOR_BORRAR)

        # Fila 1: Q-P en columnas 0-9, ENTER en columna 10 (ocupa filas 1-2)
        for i, ch in enumerate(FILA_2):
            btn = self._tecla(cont, self._texto_tecla(ch),
                               lambda c=ch: self._pulsar(c), row=1, col=i)
            self.letter_buttons[ch] = btn
        self._tecla(cont, "ENTER", self._confirmar, row=1, col=10, rowspan=2,
                    color=COLOR_ENTER)

        # Fila 2: A-Ñ en columnas 0-9 (columna 10 ya ocupada por ENTER)
        for i, ch in enumerate(FILA_3):
            btn = self._tecla(cont, self._texto_tecla(ch),
                               lambda c=ch: self._pulsar(c), row=2, col=i)
            self.letter_buttons[ch] = btn

        # Fila 3: MAYUS en columna 0, luego Z X C V B N M , . - en columnas 1-10
        self.mayus_button = self._tecla(
            cont, self._texto_mayus(), self._toggle_mayus, row=3, col=0,
            color=self._color_mayus(),
        )
        for i, ch in enumerate(FILA_4):
            if ch.isalpha():
                btn = self._tecla(cont, self._texto_tecla(ch),
                                   lambda c=ch: self._pulsar(c), row=3, col=i + 1)
                self.letter_buttons[ch] = btn
            else:
                self._tecla(cont, ch, lambda c=ch: self._pulsar(c), row=3, col=i + 1)

        # Fila 4: CERRAR (columnas 0-2) + ESPACIO (columnas 3-10)
        self._tecla(cont, "CERRAR", self.ocultar, row=4, col=0, colspan=3,
                    color=COLOR_CERRAR)
        self._tecla(cont, "ESPACIO", self._espacio, row=4, col=3, colspan=8,
                    color=COLOR_ESPACIO)

    # ------------------------------------------------------------------
    # NUMÉRICO: 4 columnas x 4 filas
    #   Fila 0: 7 8 9 BORRAR
    #   Fila 1: 4 5 6 -
    #   Fila 2: 1 2 3 ENTER
    #   Fila 3: 0 (colspan 2) , CERRAR
    # ------------------------------------------------------------------
    def _construir_numerico(self, cont):
        n_cols = 4
        n_rows = 4
        for c in range(n_cols):
            cont.grid_columnconfigure(c, weight=1, uniform="col")
        for r in range(n_rows):
            cont.grid_rowconfigure(r, weight=1)

        filas_digitos = [
            (["7", "8", "9"], "BORRAR", self._borrar, COLOR_BORRAR),
            (["4", "5", "6"], "-", lambda: self._pulsar("-"), COLOR_TECLA),
            (["1", "2", "3"], "ENTER", self._confirmar, COLOR_ENTER),
        ]
        for row, (digitos, extra_texto, extra_cmd, extra_color) in enumerate(filas_digitos):
            for col, d in enumerate(digitos):
                self._tecla(cont, d, lambda c=d: self._pulsar(c), row=row, col=col)
            self._tecla(cont, extra_texto, extra_cmd, row=row, col=3, color=extra_color)

        # Última fila: "0" ancho (colspan 2), coma, CERRAR
        self._tecla(cont, "0", lambda: self._pulsar("0"), row=3, col=0, colspan=2)
        self._tecla(cont, ",", lambda: self._pulsar(","), row=3, col=2)
        self._tecla(cont, "CERRAR", self.ocultar, row=3, col=3, color=COLOR_CERRAR)

    # ------------------------------------------------------------------
    def _texto_tecla(self, ch):
        if ch.isalpha() and self.mayus:
            return ch.upper()
        elif ch.isalpha():
            return ch.lower()
        return ch  # dígitos u otros símbolos, sin cambio

    def _texto_mayus(self):
        return "MAYÚS ▲" if self.mayus else "MAYÚS"

    def _color_mayus(self):
        return COLOR_MAYUS_ON if self.mayus else COLOR_TECLA

    def _pulsar(self, char):
        if self.target is None:
            return
        char_final = char if (self.mayus and char.isalpha()) or self.numeric_only \
            else char.lower()
        self._insertar(char_final)

    def _espacio(self):
        self._insertar(" ")

    def _borrar(self):
        if self.target is None:
            return
        if isinstance(self.target, tk.Text):
            self.target.delete("insert-1c", "insert")
        else:
            pos = self.target.index(tk.INSERT)
            if pos > 0:
                self.target.delete(pos - 1, pos)
        self._notificar_cambio()

    def _toggle_mayus(self):
        self.mayus = not self.mayus
        if self.mayus_button is not None:
            self.mayus_button.config(text=self._texto_mayus(), bg=self._color_mayus(),
                                      relief="sunken" if self.mayus else "raised")
        for ch, btn in self.letter_buttons.items():
            btn.config(text=self._texto_tecla(ch))

    def _insertar(self, texto):
        if isinstance(self.target, tk.Text):
            self.target.insert("insert", texto)
        else:
            self.target.insert(tk.INSERT, texto)
        self._notificar_cambio()

    def _notificar_cambio(self):
        """Genera un evento virtual para que los widgets que necesiten
        reaccionar al texto introducido desde el teclado virtual (que no
        pasa por eventos de teclado reales) puedan enterarse del cambio."""
        try:
            self.target.event_generate("<<VirtualKeyboardInput>>")
        except tk.TclError:
            pass


def attach_keyboard(widget, numeric_only=False):
    """
    Asocia el teclado virtual a un widget Entry o Text: se mostrará al
    pulsar/enfocar el widget.
    """
    def _on_focus(event):
        root = widget.winfo_toplevel()
        kb = VirtualKeyboard.get(root)
        kb.mostrar(widget, numeric_only=numeric_only)

    widget.bind("<Button-1>", _on_focus)
    widget.bind("<FocusIn>", _on_focus)
