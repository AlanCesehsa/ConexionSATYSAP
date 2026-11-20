"""
========================================================================================
CONCILIACIÓN FISCAL SAP vs. SAT — GRUPO CESEHSA
========================================================================================
Interfaz limpia, minimalista y directa:
- Soporte para auditar un solo mes, un rango de meses o el año acumulado.
- Corrección de bugs visuales en desplegables (no se vacían al hacer clic).
- Nombre de salida automático y claro: Conciliacion-{sociedad}-{año}-{meses}.xlsx
- Selección clara de carpeta de destino.
- Conexión directa a SQL Server SAP (192.168.1.250).
========================================================================================
"""

import os
import sys
import glob
import threading
import datetime
import warnings
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional
import pandas as pd

# Silenciar advertencias informativas de librerías
warnings.filterwarnings('ignore')

from sat_descarga_masiva import SATDescargaMasiva
from motor_conciliacion import conciliar_sociedad, exportar_excel_multi_sociedad, estructurar_sat_df
from descargar_sat import descargar_anio_completo
from conexion_sap_db import obtener_facturas_sap_db, probar_conexion_sap

# Alta resolución DPI en Windows
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass


def obtener_ruta_base() -> str:
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def obtener_carpeta_descargas_defecto() -> str:
    """Devuelve la carpeta de Descargas del usuario o reportes/."""
    descargas = os.path.join(os.path.expanduser("~"), "Downloads")
    if os.path.exists(descargas):
        return os.path.abspath(descargas)
    base = obtener_ruta_base()
    rep = os.path.join(base, "reportes")
    os.makedirs(rep, exist_ok=True)
    return rep


def buscar_archivo_sap_defecto(base_dir: str) -> str:
    """Encuentra automáticamente el archivo de facturas SAP en las ubicaciones habituales."""
    candidatos = [
        r"c:\Users\Stock CESEHSA\Downloads\EJEMPLO DE CONCILIACION.xls",
        r"c:\Users\Stock CESEHSA\Documents\Reporte_SAPvsSAT (2)\Reporte_SAPvsSAT.xlsx",
        os.path.join(base_dir, "reportes", "Reporte_SAPvsSAT.xlsx"),
        os.path.join(base_dir, "Reporte_SAPvsSAT.xlsx")
    ]
    for c in candidatos:
        if os.path.exists(c):
            return os.path.abspath(c)
    
    patrones = [
        r"c:\Users\Stock CESEHSA\Downloads\*CONCILIA*.xls*",
        r"c:\Users\Stock CESEHSA\Downloads\*SAP*.xls*",
        os.path.join(base_dir, "*SAP*.xls*")
    ]
    for p in patrones:
        encontrados = glob.glob(p)
        if encontrados:
            return os.path.abspath(encontrados[0])
    return ""


class AppConciliacionFiscal:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Conciliación Fiscal — SAP vs. SAT")
        
        # Tamaño cómodo, compacto y adaptable a cualquier resolución sin cortes
        self.root.geometry("490x580")
        self.root.minsize(460, 520)
        self.root.resizable(True, True)
        self._centrar_ventana(490, 580)

        self.base_dir = obtener_ruta_base()
        self.ruta_sap = buscar_archivo_sap_defecto(self.base_dir)
        self.carpeta_salida = obtener_carpeta_descargas_defecto()
        self.archivo_salida = None

        self._configurar_tema()
        self._construir_ui()

    def _centrar_ventana(self, ancho, alto):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - ancho) // 2)
        if sh <= 768:
            y = 20
        else:
            y = max(0, (sh - alto) // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

    def _configurar_tema(self):
        self.COLOR_BG = "#F8FAFC"
        self.COLOR_CARD = "#FFFFFF"
        self.COLOR_PRIMARY = "#2563EB"
        self.COLOR_PRIMARY_HOVER = "#1D4ED8"
        self.COLOR_TEXT_MAIN = "#0F172A"
        self.COLOR_TEXT_MUTED = "#64748B"
        self.COLOR_BORDER = "#E2E8F0"
        self.COLOR_SUCCESS = "#16A34A"

        self.root.configure(bg=self.COLOR_BG)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background=self.COLOR_BG, font=("Segoe UI", 9))
        style.configure("TCombobox", padding=4, font=("Segoe UI", 9))

        # Corrección del bug visual: mapear colores para que nunca se vuelva blanco sobre blanco
        style.map("TCombobox",
            fieldbackground=[("readonly", "#FFFFFF"), ("!disabled", "#FFFFFF")],
            foreground=[("readonly", "#0F172A"), ("!disabled", "#0F172A"), ("focus", "#0F172A")],
            selectbackground=[("readonly", "#FFFFFF"), ("focus", "#FFFFFF")],
            selectforeground=[("readonly", "#0F172A"), ("focus", "#0F172A")]
        )
        style.configure("TProgressbar", background=self.COLOR_PRIMARY, troughcolor=self.COLOR_BORDER)

    def _construir_ui(self):
        # 1. ENCABEZADO
        header = tk.Frame(self.root, bg=self.COLOR_CARD, padx=14, pady=5)
        header.pack(fill=tk.X)

        lbl_tit = tk.Label(
            header,
            text="Conciliación Fiscal",
            font=("Segoe UI", 12, "bold"),
            fg=self.COLOR_TEXT_MAIN,
            bg=self.COLOR_CARD
        )
        lbl_tit.pack(anchor=tk.W)

        lbl_sub = tk.Label(
            header,
            text="Auditoría automática de ingresos entre SAP y SAT con e.firma",
            font=("Segoe UI", 8),
            fg=self.COLOR_TEXT_MUTED,
            bg=self.COLOR_CARD
        )
        lbl_sub.pack(anchor=tk.W, pady=(1, 0))

        div = tk.Frame(self.root, bg=self.COLOR_BORDER, height=1)
        div.pack(fill=tk.X)

        # 2. CONTENEDOR PRINCIPAL
        main_body = tk.Frame(self.root, bg=self.COLOR_BG, padx=12, pady=6)
        main_body.pack(fill=tk.BOTH, expand=True)

        # TARJETA BLANCA DE CONFIGURACIÓN
        card = tk.Frame(main_body, bg=self.COLOR_CARD, padx=10, pady=6, relief="solid", bd=1)
        card.config(highlightbackground=self.COLOR_BORDER, highlightcolor=self.COLOR_BORDER)
        card.pack(fill=tk.X)

        # Campo: Sociedad
        lbl_soc = tk.Label(card, text="Sociedad", font=("Segoe UI", 8, "bold"), fg=self.COLOR_TEXT_MAIN, bg=self.COLOR_CARD)
        lbl_soc.pack(anchor=tk.W, pady=(0, 1))

        opciones_soc = [
            "Todas las sociedades (CESEHSA + CSH + YUKME)",
            "CESEHSA (CPR100602NJ2)",
            "CSH (CPR1006024T4)",
            "YUKME (YUK990113KU9)"
        ]
        self.cb_sociedad = ttk.Combobox(card, values=opciones_soc, state="readonly", font=("Segoe UI", 8))
        self.cb_sociedad.current(0)
        self.cb_sociedad.pack(fill=tk.X, pady=(0, 4))
        self.cb_sociedad.bind("<FocusIn>", lambda e: e.widget.selection_clear())

        # Campo: Periodo (Año, Mes Inicial y Mes Final)
        periodo_frame = tk.Frame(card, bg=self.COLOR_CARD)
        periodo_frame.pack(fill=tk.X, pady=(0, 2))

        # Columna 1: Año
        col_anio = tk.Frame(periodo_frame, bg=self.COLOR_CARD)
        col_anio.pack(side=tk.LEFT, fill=tk.X, expand=False, padx=(0, 3))

        lbl_a = tk.Label(col_anio, text="Año", font=("Segoe UI", 8, "bold"), fg=self.COLOR_TEXT_MAIN, bg=self.COLOR_CARD)
        lbl_a.pack(anchor=tk.W, pady=(0, 1))

        anio_actual = datetime.datetime.now().year
        anios = [str(a) for a in range(anio_actual, 2023, -1)]
        self.cb_anio = ttk.Combobox(col_anio, values=anios, state="readonly", font=("Segoe UI", 8), width=6)
        self.cb_anio.set(str(anio_actual))
        self.cb_anio.pack(fill=tk.X)
        self.cb_anio.bind("<<ComboboxSelected>>", self._actualizar_meses)
        self.cb_anio.bind("<FocusIn>", lambda e: e.widget.selection_clear())

        # Columna 2: Mes Inicial (Desde)
        col_mes_ini = tk.Frame(periodo_frame, bg=self.COLOR_CARD)
        col_mes_ini.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 2))

        lbl_m1 = tk.Label(col_mes_ini, text="Mes inicial", font=("Segoe UI", 8, "bold"), fg=self.COLOR_TEXT_MAIN, bg=self.COLOR_CARD)
        lbl_m1.pack(anchor=tk.W, pady=(0, 1))

        self.meses_nombres = [
            "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril",
            "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto",
            "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"
        ]
        self.cb_mes_ini = ttk.Combobox(col_mes_ini, values=self.meses_nombres, state="readonly", font=("Segoe UI", 8), width=14)
        self.cb_mes_ini.pack(fill=tk.X)
        self.cb_mes_ini.bind("<FocusIn>", lambda e: e.widget.selection_clear())
        self.cb_mes_ini.bind("<<ComboboxSelected>>", self._on_mes_ini_change)

        # Columna 3: Mes Final (Hasta)
        col_mes_fin = tk.Frame(periodo_frame, bg=self.COLOR_CARD)
        col_mes_fin.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(2, 0))

        lbl_m2 = tk.Label(col_mes_fin, text="Mes final", font=("Segoe UI", 8, "bold"), fg=self.COLOR_TEXT_MAIN, bg=self.COLOR_CARD)
        lbl_m2.pack(anchor=tk.W, pady=(0, 1))

        self.cb_mes_fin = ttk.Combobox(col_mes_fin, values=self.meses_nombres, state="readonly", font=("Segoe UI", 8), width=14)
        self.cb_mes_fin.pack(fill=tk.X)
        self.cb_mes_fin.bind("<FocusIn>", lambda e: e.widget.selection_clear())
        self.cb_mes_fin.bind("<<ComboboxSelected>>", self._on_mes_fin_change)

        # Checkbox para seleccionar todo el año acumulado
        quick_frame = tk.Frame(card, bg=self.COLOR_CARD)
        quick_frame.pack(fill=tk.X, pady=(2, 4))

        self.var_todo_anio = tk.BooleanVar(value=False)
        chk_todo = tk.Checkbutton(
            quick_frame,
            text="Auditar todo el año acumulado (Enero al mes actual)",
            variable=self.var_todo_anio,
            command=self._toggle_todo_anio,
            font=("Segoe UI", 8),
            fg=self.COLOR_PRIMARY,
            bg=self.COLOR_CARD,
            activebackground=self.COLOR_CARD,
            activeforeground=self.COLOR_PRIMARY
        )
        chk_todo.pack(anchor=tk.W)

        self._actualizar_meses()

        # Campo: Guardar en (Selección de carpeta)
        lbl_ruta = tk.Label(card, text="Guardar reporte en", font=("Segoe UI", 8, "bold"), fg=self.COLOR_TEXT_MAIN, bg=self.COLOR_CARD)
        lbl_ruta.pack(anchor=tk.W, pady=(0, 1))

        ruta_box = tk.Frame(card, bg="#F1F5F9", padx=6, pady=2, relief="flat")
        ruta_box.pack(fill=tk.X, pady=(0, 2))

        self.lbl_ruta_texto = tk.Label(
            ruta_box,
            text=self.carpeta_salida,
            font=("Segoe UI", 8),
            fg=self.COLOR_TEXT_MAIN,
            bg="#F1F5F9",
            anchor=tk.W
        )
        self.lbl_ruta_texto.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_cambiar_ruta = tk.Button(
            ruta_box,
            text="Cambiar...",
            font=("Segoe UI", 7, "bold"),
            bg="#E2E8F0",
            fg="#334155",
            relief="flat",
            bd=0,
            padx=5,
            pady=1,
            cursor="hand2",
            command=self._seleccionar_carpeta_salida
        )
        btn_cambiar_ruta.pack(side=tk.RIGHT)

        # 3. BOTÓN PRINCIPAL DE ACCIÓN (UBICADO ARRIBA, DIRECTO BAJO LA CONFIGURACIÓN)
        self.btn_ejecutar = tk.Button(
            main_body,
            text="Iniciar Conciliación",
            font=("Segoe UI", 10, "bold"),
            bg=self.COLOR_PRIMARY,
            fg="#FFFFFF",
            activebackground=self.COLOR_PRIMARY_HOVER,
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            pady=6,
            cursor="hand2",
            command=self._iniciar_conciliacion
        )
        self.btn_ejecutar.pack(fill=tk.X, pady=(6, 3))

        # 4. BARRA DE PROGRESO (se muestra al dar clic)
        self.progreso = ttk.Progressbar(main_body, mode="indeterminate")

        # 5. CAJA DE ESTADO
        self.card_status = tk.Frame(main_body, bg="#FFFFFF", padx=8, pady=5, relief="solid", bd=1)
        self.card_status.config(highlightbackground=self.COLOR_BORDER, highlightcolor=self.COLOR_BORDER)
        self.card_status.pack(fill=tk.X, pady=(0, 3))

        self.lbl_estado = tk.Label(
            self.card_status,
            text="Listo para conciliar. Haz clic en 'Iniciar Conciliación'.",
            font=("Segoe UI", 8),
            fg=self.COLOR_TEXT_MAIN,
            bg="#FFFFFF",
            wraplength=450,
            justify=tk.LEFT
        )
        self.lbl_estado.pack(anchor=tk.W)

        # 6. BOTÓN VERDE PARA ABRIR EXCEL (se empaca al terminar)
        self.btn_abrir = tk.Button(
            main_body,
            text="Abrir Reporte en Excel",
            font=("Segoe UI", 9, "bold"),
            bg=self.COLOR_SUCCESS,
            fg="#FFFFFF",
            activebackground="#15803D",
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            pady=5,
            cursor="hand2",
            command=self._abrir_excel
        )

        # 7. ENLACES SECUNDARIOS
        btn_solo_sat = tk.Button(
            main_body,
            text="¿Solo descargar facturas del SAT (Año completo)? Clic aquí",
            font=("Segoe UI", 7, "underline"),
            fg=self.COLOR_TEXT_MUTED,
            bg=self.COLOR_BG,
            activeforeground=self.COLOR_PRIMARY,
            activebackground=self.COLOR_BG,
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self._iniciar_descarga_anual
        )
        btn_solo_sat.pack()

        # Indicador de conexión a SAP
        self.btn_origen_sap = tk.Button(
            main_body,
            text="Fuente SAP: Base de datos SQL (192.168.1.250) [En línea]",
            font=("Segoe UI", 7),
            fg="#16A34A",
            bg=self.COLOR_BG,
            activeforeground=self.COLOR_PRIMARY,
            activebackground=self.COLOR_BG,
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self._seleccionar_archivo_sap
        )
        self.btn_origen_sap.pack(pady=(1, 0))

    def _actualizar_meses(self, event=None):
        anio_sel = int(self.cb_anio.get())
        ahora = datetime.datetime.now()
        if anio_sel == ahora.year:
            meses_validos = self.meses_nombres[:ahora.month]
        else:
            meses_validos = self.meses_nombres

        self.cb_mes_ini['values'] = meses_validos
        self.cb_mes_fin['values'] = meses_validos

        # Mes actual por defecto
        mes_idx = min(ahora.month - 1, len(meses_validos) - 1)
        if not self.var_todo_anio.get():
            self.cb_mes_ini.current(mes_idx)
            self.cb_mes_fin.current(mes_idx)
        else:
            self.cb_mes_ini.current(0)
            self.cb_mes_fin.current(mes_idx)

    def _on_mes_ini_change(self, event=None):
        idx_ini = self.cb_mes_ini.current()
        idx_fin = self.cb_mes_fin.current()
        if idx_ini > idx_fin:
            self.cb_mes_fin.current(idx_ini)
        self.var_todo_anio.set(False)

    def _on_mes_fin_change(self, event=None):
        idx_ini = self.cb_mes_ini.current()
        idx_fin = self.cb_mes_fin.current()
        if idx_fin < idx_ini:
            self.cb_mes_ini.current(idx_fin)
        self.var_todo_anio.set(False)

    def _toggle_todo_anio(self):
        if self.var_todo_anio.get():
            self.cb_mes_ini.current(0)  # Enero
            valores = self.cb_mes_fin['values']
            self.cb_mes_fin.current(len(valores) - 1)
        else:
            ahora = datetime.datetime.now()
            valores = self.cb_mes_fin['values']
            mes_idx = min(ahora.month - 1, len(valores) - 1)
            self.cb_mes_ini.current(mes_idx)
            self.cb_mes_fin.current(mes_idx)

    def _seleccionar_carpeta_salida(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta donde guardar el reporte")
        if carpeta:
            self.carpeta_salida = carpeta
            self.lbl_ruta_texto.config(text=carpeta)

    def _seleccionar_archivo_sap(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo de facturas SAP",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if ruta:
            self.ruta_sap = ruta
            self.btn_origen_sap.config(text=f"Fuente SAP: {os.path.basename(ruta)}  (clic para cambiar archivo)")

    def _obtener_nombre_abreviado(self, sociedad_texto: str) -> str:
        if "Todas" in sociedad_texto:
            return "TODAS"
        if "CESEHSA" in sociedad_texto:
            return "CESEHSA"
        if "CSH" in sociedad_texto:
            return "CSH"
        if "YUKME" in sociedad_texto:
            return "YUKME"
        return "SOCIEDAD"

    def _set_estado(self, texto: str, es_error: bool = False, es_exito: bool = False):
        color = self.COLOR_TEXT_MAIN
        if es_error:
            color = "#DC2626"
        elif es_exito:
            color = self.COLOR_SUCCESS

        def update():
            try:
                self.lbl_estado.config(text=texto, fg=color)
                # Actualizar el texto del botón cuando esté ocupado
                if not es_exito and not es_error:
                    t_low = texto.lower()
                    if any(w in t_low for w in ["excel", "escribiendo", "formateando", "diseñando", "guardando", "construcción"]):
                        self.btn_ejecutar.config(text="Construyendo Excel... Por favor espere", bg="#D97706")
                    elif any(w in t_low for w in ["sat", "descargando", "autenticando", "solicitando"]):
                        self.btn_ejecutar.config(text="Descargando del SAT...", bg="#2563EB")
                    elif any(w in t_low for w in ["sap", "consultando"]):
                        self.btn_ejecutar.config(text="Consultando SAP...", bg="#2563EB")
                    elif any(w in t_low for w in ["conciliando"]):
                        self.btn_ejecutar.config(text="Conciliando facturas...", bg="#2563EB")
            except Exception:
                pass
        self.root.after(0, update)

    def _iniciar_conciliacion(self):
        self.btn_ejecutar.config(state=tk.DISABLED, text="Iniciando...", bg="#94A3B8")
        self.progreso.pack(fill=tk.X, pady=(0, 4), before=self.card_status)
        self.progreso.start(10)
        self.btn_abrir.pack_forget()
        self._set_estado("Iniciando conciliación...")

        hilo = threading.Thread(target=self._hilo_conciliacion, daemon=True)
        hilo.start()

    def _iniciar_descarga_anual(self):
        self.btn_ejecutar.config(state=tk.DISABLED, text="Iniciando descarga SAT...", bg="#94A3B8")
        self.progreso.pack(fill=tk.X, pady=(0, 4), before=self.card_status)
        self.progreso.start(10)
        self.btn_abrir.pack_forget()
        self._set_estado("Iniciando descarga anual del SAT...")

        hilo = threading.Thread(target=self._hilo_descarga_anual, daemon=True)
        hilo.start()

    def _hilo_descarga_anual(self):
        try:
            soc_sel = self.cb_sociedad.get()
            soc_abrev = self._obtener_nombre_abreviado(soc_sel)
            if soc_abrev == "TODAS":
                sociedades = ["CESEHSA", "CSH", "YUKME"]
            else:
                sociedades = [soc_abrev]

            anio = int(self.cb_anio.get())
            dfs_sat = {}

            for soc in sociedades:
                self._set_estado(f"Descargando comprobantes {anio} del SAT para {soc}...")
                df = descargar_anio_completo(soc, anio, carpeta_salida=self.carpeta_salida, abrir_excel=False)
                if df is not None and not df.empty:
                    dfs_sat[soc] = df

            if not dfs_sat:
                self._set_estado("El SAT no devolvió registros para este año.", es_error=True)
                return

            nombre_archivo = f"DescargaSAT-{soc_abrev}-{anio}.xlsx"
            ruta_final = os.path.join(self.carpeta_salida, nombre_archivo)

            if len(sociedades) == 1:
                soc = sociedades[0]
                df_fmt = estructurar_sat_df(dfs_sat[soc])
                with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
                    df_fmt.to_excel(writer, sheet_name=f"SAT ({soc})", index=False)
            else:
                with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
                    for soc, df in dfs_sat.items():
                        df_fmt = estructurar_sat_df(df)
                        df_fmt.to_excel(writer, sheet_name=f"SAT ({soc})", index=False)

            self.archivo_salida = ruta_final
            self._set_estado(f"Descarga finalizada con éxito:\n{nombre_archivo}", es_exito=True)
            self.root.after(0, lambda: self.btn_abrir.pack(fill=tk.X, pady=(0, 3), after=self.card_status))
            self._abrir_excel()

        except Exception as e:
            traceback.print_exc()
            self._set_estado(f"Error: {str(e)}", es_error=True)
        finally:
            def restore_ui():
                try:
                    self.progreso.stop()
                    self.progreso.pack_forget()
                    self.btn_ejecutar.config(state=tk.NORMAL, text="Iniciar Conciliación", bg=self.COLOR_PRIMARY)
                except Exception:
                    pass
            self.root.after(0, restore_ui)

    def _hilo_conciliacion(self):
        try:
            soc_sel = self.cb_sociedad.get()
            soc_abrev = self._obtener_nombre_abreviado(soc_sel)
            if soc_abrev == "TODAS":
                sociedades = ["CESEHSA", "CSH", "YUKME"]
            else:
                sociedades = [soc_abrev]

            anio = int(self.cb_anio.get())
            mes_ini = int(self.cb_mes_ini.get()[:2])
            mes_fin = int(self.cb_mes_fin.get()[:2])
            if mes_ini > mes_fin:
                mes_ini, mes_fin = mes_fin, mes_ini

            ahora = datetime.datetime.now()
            fecha_ini = datetime.datetime(anio, mes_ini, 1, 0, 0, 0)
            if mes_fin in [1, 3, 5, 7, 8, 10, 12]:
                ultimo_dia = 31
            elif mes_fin in [4, 6, 9, 11]:
                ultimo_dia = 30
            else:
                ultimo_dia = 29 if anio % 4 == 0 else 28
            fecha_fin = datetime.datetime(anio, mes_fin, ultimo_dia, 23, 59, 59)
            if fecha_fin > ahora:
                fecha_fin = ahora - datetime.timedelta(minutes=1)

            # Etiqueta de periodo para mensajes
            if mes_ini == mes_fin:
                periodo_texto = f"{anio}-{mes_ini:02d}"
                nombre_archivo = f"Conciliacion-{soc_abrev}-{anio}-{mes_ini:02d}.xlsx"
            else:
                periodo_texto = f"{anio} (Meses {mes_ini:02d} a {mes_fin:02d})"
                nombre_archivo = f"Conciliacion-{soc_abrev}-{anio}-{mes_ini:02d}-a-{mes_fin:02d}.xlsx"

            resultados_multi = {}
            total_conciliadas = 0
            total_alertas = 0

            for soc in sociedades:
                # 1. Obtener datos de SAP (Directo de la Base de Datos SQL 192.168.1.250 o Excel)
                df_sap = None
                try:
                    self._set_estado(f"Consultando facturas en servidor SAP para {soc} ({periodo_texto})...")
                    df_sap = obtener_facturas_sap_db(soc, anio, mes_ini, mes_fin)
                except Exception as e_sql:
                    # Fallback si no está en la red local
                    if self.ruta_sap and os.path.exists(self.ruta_sap):
                        try:
                            if self.ruta_sap.lower().endswith(".xls"):
                                xl_sap = pd.ExcelFile(self.ruta_sap, engine="xlrd")
                            else:
                                xl_sap = pd.ExcelFile(self.ruta_sap, engine="openpyxl")
                            hoja = next((h for h in [f"SAP_{soc}", soc, "facturas sap"] if h in xl_sap.sheet_names), xl_sap.sheet_names[0])
                            df_sap = xl_sap.parse(hoja)
                        except Exception:
                            df_sap = pd.DataFrame()
                    else:
                        raise e_sql

                if df_sap is None:
                    df_sap = pd.DataFrame()

                # 2. Descargar facturas del SAT con e.firma
                self._set_estado(f"Conectando al SAT con e.firma de {soc}...")
                sat = SATDescargaMasiva(soc, base_cert_dir=os.path.join(self.base_dir, "certificados"))
                sat.autenticar()

                self._set_estado(f"Descargando comprobantes de {soc} del SAT ({periodo_texto})...")
                id_sol, res_sat = sat.solicitar_descarga_metadatos(fecha_ini, fecha_fin)

                if id_sol:
                    df_sat = sat.verificar_y_descargar(id_sol, max_intentos=15, espera_segundos=4)
                else:
                    df_sat = pd.DataFrame()

                # 3. Conciliar
                self._set_estado(f"Conciliando facturas por UUID ({soc})...")
                res_soc = conciliar_sociedad(df_sap, df_sat, nombre_sociedad=soc, anio=anio, mes=mes_ini, mes_fin=mes_fin)
                resultados_multi[soc] = res_soc

                total_conciliadas += len(res_soc["CONCILIADAS_OK"])
                total_alertas += len(res_soc["CANCELADAS_SAT_EN_SAP"]) + len(res_soc["SAP_FALTANTES_EN_SAT"])

            ruta_salida = os.path.join(self.carpeta_salida, nombre_archivo)
            self._set_estado(f"Iniciando construcción de libro Excel: {nombre_archivo}...")
            self.root.after(0, lambda: self.btn_ejecutar.config(text="Construyendo Excel... Por favor espere", bg="#D97706"))
            ruta_salida = exportar_excel_multi_sociedad(resultados_multi, ruta_salida, callback_progreso=self._set_estado)
            self.archivo_salida = ruta_salida
            nombre_archivo = os.path.basename(ruta_salida)

            resumen_txt = (
                f"Conciliación finalizada ({periodo_texto}): {total_conciliadas} OK  |  {total_alertas} Alertas fiscales.\n"
                f"Archivo: {nombre_archivo}"
            )
            self._set_estado(resumen_txt, es_exito=True)
            self.root.after(0, lambda: self.btn_abrir.pack(fill=tk.X, pady=(0, 3), after=self.card_status))
            self._abrir_excel()

        except Exception as e:
            traceback.print_exc()
            self._set_estado(f"Error: {str(e)}", es_error=True)
        finally:
            def restore_ui():
                try:
                    self.progreso.stop()
                    self.progreso.pack_forget()
                    self.btn_ejecutar.config(state=tk.NORMAL, text="Iniciar Conciliación", bg=self.COLOR_PRIMARY)
                except Exception:
                    pass
            self.root.after(0, restore_ui)

    def _abrir_excel(self):
        if self.archivo_salida and os.path.exists(self.archivo_salida):
            try:
                os.startfile(self.archivo_salida)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")


def main():
    root = tk.Tk()
    app = AppConciliacionFiscal(root)
    root.mainloop()


if __name__ == "__main__":
    main()
