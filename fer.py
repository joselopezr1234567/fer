import os
# Forzar entorno seguro para hilos numéricos
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["VECLIB_MAXIMUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"

import threading
import time
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import re

try:
    import pyautogui
    pyautogui.FAILSAFE = True
    PYAUTOGUI_DISPONIBLE = True
except ImportError:
    PYAUTOGUI_DISPONIBLE = False

class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Filtro y Automatización Inteligente - Secuencia Exacta")
        self.root.geometry("950x680")
        self.root.minsize(800, 580)
        
        self.file_path = None
        self.df = None
        self.filtered_results = []
        self.is_running = False
        self.report_data = []

        self.create_widgets()

    def create_widgets(self):
        top_frame = ttk.LabelFrame(self.root, text=" 1. Archivo y Filtrado ", padding=12)
        top_frame.pack(fill="x", padx=15, pady=10)

        top_frame.columnconfigure(1, weight=1)

        self.btn_upload = ttk.Button(top_frame, text="Seleccionar Archivo", command=self.load_file_thread)
        self.btn_upload.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.lbl_file = ttk.Label(top_frame, text="Ningún archivo seleccionado", foreground="gray", anchor="w")
        self.lbl_file.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        self.btn_process = ttk.Button(top_frame, text="Procesar y Filtrar", command=self.process_data_thread, state="disabled")
        self.btn_process.grid(row=0, column=2, padx=5, pady=5, sticky="e")

        auto_frame = ttk.LabelFrame(self.root, text=" 2. Automatización Web (Parada: botón o mueve el mouse a una esquina) ", padding=12)
        auto_frame.pack(fill="x", padx=15, pady=5)

        auto_frame.columnconfigure(0, weight=1)
        auto_frame.columnconfigure(1, weight=1)

        self.btn_auto = ttk.Button(auto_frame, text="▶ Ejecutar Automatización Inteligente", command=self.run_automation_thread, state="disabled")
        self.btn_auto.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.btn_stop = ttk.Button(auto_frame, text="⏹ Detener Proceso", command=self.stop_automation, state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=100, mode="indeterminate")
        self.progress.pack(fill="x", padx=15, pady=5)

        bottom_frame = ttk.LabelFrame(self.root, text=" Asociaciones Encontradas (SuppName, OptSrv, location) ", padding=12)
        bottom_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        self.text_output = tk.Text(bottom_frame, wrap="none", font=("Courier New", 10))
        
        # Configurar etiquetas de colores para la interfaz
        self.text_output.tag_config("rojo", background="#ffcccc", foreground="black")      # No se oculta (hay reservas futuras)
        self.text_output.tag_config("amarillo", background="#ffffe6", foreground="black")   # Se oculta o no encontrado
        self.text_output.tag_config("normal", background="white", foreground="black")

        scrollbar_y = ttk.Scrollbar(bottom_frame, orient="vertical", command=self.text_output.yview)
        scrollbar_x = ttk.Scrollbar(bottom_frame, orient="horizontal", command=self.text_output.xview)
        
        self.text_output.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.text_output.pack(side="left", fill="both", expand=True)

    def format_path(self, path, max_len=45):
        if not path:
            return "Ningún archivo seleccionado"
        if len(path) <= max_len:
            return path
        head, tail = os.path.split(path)
        return f".../{tail}" if len(tail) < max_len - 4 else f"...{path[-(max_len-3):]}"

    def load_file_thread(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel o CSV",
            filetypes=[
                ("Archivos Compatibles", "*.xlsx *.xls *.csv"),
                ("Archivos CSV", "*.csv"),
                ("Archivos de Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*")
            ]
        )
        if not file_path:
            return

        self.file_path = file_path
        short_name = self.format_path(file_path)
        self.lbl_file.config(text=f"Cargando: {short_name}...", foreground="blue")
        self.btn_upload.config(state="disabled")
        self.progress.start(10)

        threading.Thread(target=self._load_file_task, daemon=True).start()

    def _load_file_task(self):
        try:
            if self.file_path.lower().endswith('.csv'):
                try:
                    self.df = pd.read_csv(self.file_path, sep=None, engine='python', encoding='utf-8', on_bad_lines='skip')
                except UnicodeDecodeError:
                    self.df = pd.read_csv(self.file_path, sep=None, engine='python', encoding='latin1', on_bad_lines='skip')
            else:
                try:
                    self.df = pd.read_excel(self.file_path, engine='calamine')
                except Exception:
                    self.df = pd.read_excel(self.file_path, engine='openpyxl')

            self.root.after(0, self._load_file_success)
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: self._load_file_error(err_msg))

    def _load_file_success(self):
        self.progress.stop()
        self.btn_upload.config(state="normal")
        self.btn_process.config(state="normal")
        self.lbl_file.config(text=self.format_path(self.file_path), foreground="black")
        
        total_rows = len(self.df)
        messagebox.showinfo("Éxito", f"¡Archivo cargado con éxito!\n\nFilas totales: {total_rows:,}\nHaz clic en 'Procesar y Filtrar'.")

    def _load_file_error(self, error_msg):
        self.progress.stop()
        self.btn_upload.config(state="normal")
        self.btn_process.config(state="disabled")
        self.lbl_file.config(text="Error al cargar archivo", foreground="red")
        messagebox.showerror("Error de Lectura", f"No se pudo cargar el archivo:\n{error_msg}")

    def process_data_thread(self):
        if self.df is None:
            return

        self.btn_process.config(state="disabled")
        self.btn_auto.config(state="disabled")
        self.progress.start(10)
        self.text_output.delete("1.0", tk.END)
        self.text_output.insert(tk.END, "Procesando y cargando todos los datos...\n")

        threading.Thread(target=self._process_data_task, daemon=True).start()

    def _process_data_task(self):
        try:
            self.df.columns = self.df.columns.astype(str).str.strip()
            
            col_supp = next((c for c in self.df.columns if c.lower() == "suppname"), None)
            col_opt = next((c for c in self.df.columns if c.lower() == "optsrv"), None)
            col_loc = next((c for c in self.df.columns if c.lower() == "location"), None)

            if not col_supp or not col_opt or not col_loc:
                missing = []
                if not col_supp: missing.append("SuppName")
                if not col_opt: missing.append("OptSrv")
                if not col_loc: missing.append("location")
                
                error_detalle = f"No se encontraron las columnas: {', '.join(missing)}."
                self.root.after(0, lambda: self._process_error(error_detalle))
                return

            df_clean = self.df.dropna(subset=[col_supp, col_opt, col_loc]).copy()

            if df_clean.empty:
                result_text = "No se encontraron registros válidos en el archivo.\n"
                self.filtered_results = []
            else:
                result_df = df_clean[[col_supp, col_opt, col_loc]].drop_duplicates()
                
                self.filtered_results = []
                header_text = f"{'PROVEEDOR':<30} | {'CÓDIGO (OptSrv)':<20} | {'UBICACIÓN (location)':<20}\n"
                separator_text = "-" * 76 + "\n"
                
                lines = [header_text, separator_text]
                for _, row in result_df.iterrows():
                    supp = str(row[col_supp])
                    opt = str(row[col_opt])
                    loc = str(row[col_loc])
                    self.filtered_results.append((supp, opt, loc))
                    lines.append(f"{supp:<30} | {opt:<20} | {loc:<20}\n")
                
                lines.append(f"\n[Proceso completado. Se cargaron {len(self.filtered_results):,} registros únicos.]\n")
                result_text = "".join(lines)

            self.root.after(0, lambda: self._process_success(result_text))

        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: self._process_error(err_msg))

    def _process_success(self, text):
        self.progress.stop()
        self.btn_process.config(state="normal")
        if len(self.filtered_results) > 0 and PYAUTOGUI_DISPONIBLE:
            self.btn_auto.config(state="normal")
        self.text_output.delete("1.0", tk.END)
        self.text_output.insert(tk.END, text)

    def _process_error(self, error_msg):
        self.progress.stop()
        self.btn_process.config(state="normal")
        messagebox.showerror("Error de Procesamiento", error_msg)

    def run_automation_thread(self):
        if not self.filtered_results:
            messagebox.showwarning("Aviso", "No hay registros para automatizar.")
            return

        respuesta = messagebox.askyesno(
            "Confirmar Automatización", 
            f"Se automatizarán {len(self.filtered_results)} registros totales.\n¿Deseas iniciar?"
        )
        if not respuesta:
            return

        self.report_data = []
        self.btn_auto.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_process.config(state="disabled")
        self.progress.start(10)
        self.is_running = True

        threading.Thread(target=self._automation_task, daemon=True).start()

    def stop_automation(self):
        self.is_running = False

    def _append_output_log(self, text, tag="normal"):
        def _update():
            self.text_output.insert(tk.END, text + "\n", tag)
            self.text_output.see(tk.END)
        self.root.after(0, _update)

    def _leer_texto_region(self, box, whitelist='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789/-.'):
        left, top, width, height = box
        temp_img_path = "/tmp/region_check.png"
        try:
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)
            
            cmd = f"screencapture -x -R{left},{top},{width},{height} {temp_img_path}"
            subprocess.run(cmd, shell=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

            if os.path.exists(temp_img_path):
                if os.path.getsize(temp_img_path) < 800:
                    return ""
                
                import pytesseract
                from PIL import Image
                img = Image.open(temp_img_path)
                config_ocr = f'--psm 7 -c tessedit_char_whitelist={whitelist}'
                texto = pytesseract.image_to_string(img, config=config_ocr).strip()
                return texto
        except Exception:
            pass
        return ""

    def _automation_task(self):
        try:
            pyautogui.PAUSE = 0.5

            regiones_y_centros = [
                ((338, 341, 100, 33), (388, 357.5)),  # 1
                ((338, 374, 100, 33), (388, 390.5)),  # 2
                ((338, 407, 100, 33), (388, 423.5)),  # 3
                ((338, 440, 100, 33), (388, 456.5)),  # 4
                ((338, 473, 100, 33), (388, 489.5)),  # 5
                ((338, 506, 100, 33), (388, 522.5)),  # 6
                ((338, 539, 100, 33), (388, 555.5)),  # 7
                ((338, 572, 100, 33), (388, 588.5)),  # 8
                ((338, 605, 100, 33), (388, 621.5)),  # 9
                ((338, 638, 100, 33), (388, 654.5)),  # 10
                ((338, 671, 100, 33), (388, 687.5)),  # 11
                ((338, 704, 100, 33), (388, 720.5)),  # 12
                ((338, 737, 100, 33), (388, 753.5)),  # 13
                ((338, 770, 100, 33), (388, 786.5)),  # 14
                ((338, 803, 100, 33), (388, 819.5)),  # 15
                ((338, 836, 100, 33), (388, 852.5)),  # 16
                ((338, 869, 100, 33), (388, 885.5)),  # 17
                ((338, 902, 100, 33), (388, 918.5)),  # 18
                ((338, 935, 100, 33), (388, 951.5)),  # 19
                ((338, 968, 100, 20), (388, 978.0))   # 20
            ]

            # Generar las 9 regiones de fechas basadas en X=1171, Y=382 hasta Y=1009 separadas por 33 hacia abajo
            regiones_fechas = []
            y_inicial = 382
            x_inicial = 1171
            w_fecha = 126
            h_fecha = 33
            for f_idx in range(9):
                current_y = y_inicial + (f_idx * 33)
                regiones_fechas.append((x_inicial, current_y, w_fecha, h_fecha))

            if self.is_running:
                pyautogui.click(298, 1090)
                time.sleep(0.5)

            for i, (supp, opt_srv, target_location) in enumerate(self.filtered_results):
                if not self.is_running:
                    break

                supp_lower = str(supp).strip().lower()
                es_mitta_sixt = supp_lower in ["mitta", "sixt rent a car (ex europcar)"]

                # 1. Clic lupa producto (X=1060, Y=305)
                pyautogui.click(1060, 305)
                time.sleep(0.4)
                if not self.is_running: break

                # 3. Clic campo código (X=786, Y=449) y escribir código
                pyautogui.click(786, 449)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')
                pyautogui.write(str(opt_srv), interval=0.03)
                time.sleep(0.4)
                if not self.is_running: break

                # 4. Clic Results (X=479, Y=278)
                pyautogui.click(479, 278)
                time.sleep(1.2)
                if not self.is_running: break

                # 5 y 6. Leer regiones de localización y hacer clic si coincide
                match_encontrado = False
                for idx, (box, center) in enumerate(regiones_y_centros):
                    if not self.is_running: break
                    texto_leido = self._leer_texto_region(box, whitelist='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/-')
                    texto_leido = texto_leido.replace('1', 'I').replace('0', 'O').replace('5', 'S').replace('8', 'B')

                    if not texto_leido or texto_leido == "" or texto_leido == "OO":
                        break

                    target_clean = str(target_location).strip().upper()
                    leido_clean = str(texto_leido).strip().upper()

                    if target_clean in leido_clean or leido_clean in target_clean:
                        pyautogui.click(center[0], center[1])
                        time.sleep(0.4)
                        match_encontrado = True
                        break

                if not match_encontrado and self.is_running:
                    pyautogui.click(1295, 216)
                    time.sleep(0.4)
                    # No encontrado / Ya ocultado anteriormente -> Registrar y marcar en amarillo
                    msg = f"[AMARILLO] Código no encontrado con ubicación (Ya ocultado/omitido): {supp} | {opt_srv} | {target_location}"
                    self._append_output_log(msg, "amarillo")
                    self.report_data.append({
                        'Proveedor': supp,
                        'Codigo': opt_srv,
                        'Ubicacion': target_location,
                        'Estado': 'No encontrado / Ya ocultado'
                    })
                    continue

                if not self.is_running: break

                if es_mitta_sixt:
                    # --- FLUJO ESPECÍFICO MITTA Y SIXT ---
                    pyautogui.click(1162, 559)
                    time.sleep(0.4)

                    if not self.is_running: break

                    pyautogui.click(1554, 364)
                    time.sleep(0.4)

                    if not self.is_running: break

                    pyautogui.click(1520, 226)
                    time.sleep(0.3)
                    pyautogui.click(1520, 226)
                    time.sleep(0.8)

                    self.report_data.append({
                        'Proveedor': supp,
                        'Codigo': opt_srv,
                        'Ubicacion': target_location,
                        'Estado': 'Procesado Mitta/Sixt'
                    })
                
                else:
                    # --- FLUJO UNIFICADO (SEGÚN DIAGRAMA DE FLUJO) ---
                    # 7. Clic X=334, Y=649
                    pyautogui.click(334, 649)
                    time.sleep(0.4)
                    if not self.is_running: break

                    # 8. Clic X=356, Y=544
                    pyautogui.click(356, 544)
                    time.sleep(0.4)
                    if not self.is_running: break

                    # 9. ¿Está oculto? (Doble clic en X=1222, Y=365)
                    pyautogui.click(1222, 365)
                    time.sleep(0.3)
                    pyautogui.click(1222, 365)
                    time.sleep(0.8)
                    if not self.is_running: break

                    # 10. Revisar si hay reservas próximas (leer regiones de fechas) con LOGS y Soporte de Meses
                    hoy = datetime.now().date()
                    hay_posterior = False

                    print(f"\n--- [LOG] Analizando fechas para Proveedor: {supp} | Código: {opt_srv} ---")
                    print(f"[LOG] Fecha de referencia (Hoy): {hoy}")

                    meses_map = {
                        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
                        'ene': 1, 'abr': 4, 'ago': 8, 'dic': 12
                    }

                    for r_idx, r_box in enumerate(regiones_fechas):
                        if not self.is_running: break
                        fecha_txt = self._leer_texto_region(r_box)
                        
                        if not fecha_txt:
                            print(f"[LOG] Región Fecha #{r_idx+1}: (Vacía / No se detectó texto)")
                            continue

                        print(f"[LOG] Región Fecha #{r_idx+1}: OCR leyó -> '{fecha_txt}'")
                        parsed_ok = False
                        dt_parsed = None

                        match_date = re.search(r'(\d{1,2})[-/]?([A-Za-z]{3})[-/]?(\d{4})', fecha_txt)
                        if match_date:
                            d, m_str, y = match_date.groups()
                            m_lower = m_str.lower()
                            if m_lower in meses_map:
                                try:
                                    dt_parsed = datetime(int(y), meses_map[m_lower], int(d)).date()
                                    parsed_ok = True
                                except ValueError:
                                    pass

                        if not parsed_ok:
                            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
                                try:
                                    clean_num = fecha_txt.strip().split()[-1]
                                    dt_parsed = datetime.strptime(clean_num, fmt).date()
                                    parsed_ok = True
                                    break
                                except ValueError:
                                    continue

                        if parsed_ok:
                            es_posterior = dt_parsed > hoy
                            print(f"[LOG]   -> Parseado con éxito: {dt_parsed} | ¿Es posterior a hoy?: {es_posterior}")
                            if es_posterior:
                                hay_posterior = True
                                break
                        else:
                            print(f"[LOG]   -> ADVERTENCIA: El texto '{fecha_txt}' no se pudo convertir a formato de fecha.")

                        if hay_posterior:
                            break

                    print(f"[LOG] Resultado final de fechas para este registro -> ¿Hay fecha posterior?: {hay_posterior}")
                    print("-" * 70)

                    if hay_posterior:
                        # Si hay próximas (True) -> Marcar en ROJO, empezar desde la coordenada 4 (Results X=479, Y=278) y ver el siguiente código
                        msg = f"[ROJO] No se oculta (Tiene reservas próximas): {supp} | {opt_srv} | {target_location}"
                        self._append_output_log(msg, "rojo")
                        self.report_data.append({
                            'Proveedor': supp,
                            'Codigo': opt_srv,
                            'Ubicacion': target_location,
                            'Estado': 'No se oculta (Reservas futuras)'
                        })

                        pyautogui.click(1295, 216)
                        time.sleep(0.4)
                        if not self.is_running: break

                        pyautogui.click(786, 449)
                        pyautogui.hotkey('ctrl', 'a')
                        pyautogui.press('backspace')
                        pyautogui.write(str(opt_srv), interval=0.03)
                        time.sleep(0.4)
                        if not self.is_running: break

                        # Coordenada 4 (Results)
                        pyautogui.click(479, 278)
                        time.sleep(1.2)
                        if not self.is_running: break

                        match_nuevamente = False
                        for idx, (box, center) in enumerate(regiones_y_centros):
                            if not self.is_running: break
                            texto_leido = self._leer_texto_region(box, whitelist='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/-')
                            texto_leido = texto_leido.replace('1', 'I').replace('0', 'O').replace('5', 'S').replace('8', 'B')

                            if not texto_leido or texto_leido == "" or texto_leido == "OO":
                                break

                            target_clean = str(target_location).strip().upper()
                            leido_clean = str(texto_leido).strip().upper()

                            if target_clean in leido_clean or leido_clean in target_clean:
                                pyautogui.click(center[0], center[1])
                                time.sleep(0.4)
                                match_nuevamente = True
                                break

                        if not match_nuevamente and self.is_running:
                            pyautogui.click(1295, 216)
                            time.sleep(0.4)
                    else:
                        # Si no hay próximas / False (se oculta) -> Marcar en AMARILLO y secuencia de ocultamiento solicitada
                        msg = f"[AMARILLO] Se oculta (Sin reservas futuras): {supp} | {opt_srv} | {target_location}"
                        self._append_output_log(msg, "amarillo")
                        self.report_data.append({
                            'Proveedor': supp,
                            'Codigo': opt_srv,
                            'Ubicacion': target_location,
                            'Estado': 'Ocultado'
                        })

                        # 1. Clic X=395, Y=226
                        pyautogui.click(395, 226)
                        time.sleep(0.4)
                        if not self.is_running: break

                        # 2. Clic X=380, Y=279
                        pyautogui.click(380, 279)
                        time.sleep(0.4)
                        if not self.is_running: break

                        # 3. Clic X=1269, Y=559
                        pyautogui.click(1269, 559)
                        time.sleep(0.4)
                        if not self.is_running: break

                        # 4. Clic X=1563, Y=363
                        pyautogui.click(1563, 363)
                        time.sleep(0.4)
                        if not self.is_running: break

                        # 5. Dos clics en X=1520, Y=226 separados por 0.3 segundos (Exit product y exit supplier)
                        pyautogui.click(1520, 226)
                        time.sleep(0.3)
                        pyautogui.click(1520, 226)
                        time.sleep(0.8)

            # Generar archivo CSV al terminar
            if self.report_data:
                try:
                    df_report = pd.DataFrame(self.report_data)
                    csv_filename = "reporte_automatizacion.csv"
                    df_report.to_csv(csv_filename, index=False, encoding='utf-8-sig')
                    print(f"\n[LOG] Reporte CSV guardado exitosamente en: {os.path.abspath(csv_filename)}")
                except Exception as ex:
                    print(f"\n[LOG] Error al guardar archivo CSV: {ex}")

            if self.is_running:
                self.root.after(0, self._automation_success)
            else:
                self.root.after(0, self._automation_stopped)

        except pyautogui.FailSafeException:
            self.is_running = False
            self.root.after(0, self._automation_failsafe)
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: self._process_error(f"Error en automatización: {err_msg}"))

    def _automation_success(self):
        self.progress.stop()
        self.btn_process.config(state="normal")
        self.btn_auto.config(state="normal")
        self.btn_stop.config(state="disabled")
        messagebox.showinfo("Automatización Finalizada", "¡Se procesaron todos los registros y se generó el archivo 'reporte_automatizacion.csv'!")

    def _automation_stopped(self):
        self.progress.stop()
        self.btn_process.config(state="normal")
        self.btn_auto.config(state="normal")
        self.btn_stop.config(state="disabled")
        messagebox.showwarning("Automatización Detenida", "El proceso fue detenido por el usuario.")

    def _automation_failsafe(self):
        self.progress.stop()
        self.btn_process.config(state="normal")
        self.btn_auto.config(state="normal")
        self.btn_stop.config(state="disabled")
        messagebox.showerror("Parada de Emergencia", "¡Se activó la parada de emergencia porque el mouse llegó a la esquina de la pantalla!")

    def _process_error(self, error_msg):
        self.progress.stop()
        self.btn_process.config(state="normal")
        self.btn_stop.config(state="disabled")
        messagebox.showerror("Error de Automatización", error_msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()