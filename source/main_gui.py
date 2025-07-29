import tkinter as tk
from tkinter import filedialog, messagebox
import os
from openpyxl import load_workbook
from processors import df_to_wb, file_to_df, launch_pdf_gui  # Asumiendo que estos están implementados
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Documents Merger")
        self.root.geometry("700x700")

        self.file_inputs = {}
        self.frames = {}
        self.data_frames = {}

        self.doc_labels = [
            "Cuenta BCI 18",
            "Cuenta BCI 85",
            "Transferencias",
            "Cuenta BCI Comercio Exterior",
            "Tarjeta de crédito nacional 24",
            "Tarjeta de crédito nacional 69",
            "Tarjeta de crédito internacional",
            "Transbank",
            "Siteminder",
            "Banco Security",
        ]

        self.sheet_to_labels = {
            "BCI": [
                "Cuenta BCI 85",
                "Transferencias",
                "Cuenta BCI Comercio Exterior",
                "Tarjeta de crédito nacional 24",
                "Tarjeta de crédito nacional 69",
                "Tarjeta de crédito internacional",
            ],
            "BCI FondRendir": ["Cuenta BCI 18"],
            "Transbank": ["Transbank"],
            "Siteminder": ["Siteminder"],
            "Security": ["Banco Security"],
        }

        self.status_labels = {}
        self.output_dir = tk.StringVar(value=os.path.abspath("outputs"))
        self.template_path = tk.StringVar(value=os.path.abspath("template_EERR.xlsx"))
        self.model_dir = tk.StringVar(value=os.path.abspath("models")) 
        self._add_path_selector("Carpeta de salida", self.output_dir, 0)
        self._add_path_selector("Plantilla Excel", self.template_path, 1, file=True)
        self._add_path_selector("Modelo de predicción", self.model_dir, 2)
        # Separator between directory selectors and file inputs
        separator = tk.Frame(self.root, height=2, bd=1, relief="sunken", bg="gray")
        separator.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 10))

        title_row = 3
        tk.Label(self.root, text="Documento", font=("Helvetica", 14, "bold")).grid(row=title_row, column=0, padx=10, sticky="w")
        tk.Label(self.root, text="Cargar", font=("Helvetica", 14, "bold")).grid(row=title_row, column=1, padx=5, sticky="ew")
        tk.Label(self.root, text="Cargar DataFrame", font=("Helvetica", 14, "bold")).grid(row=title_row, column=2, padx=5, sticky="ew")
        tk.Label(self.root, text="Guardar DataFrame", font=("Helvetica", 14, "bold")).grid(row=title_row, column=3, padx=5, sticky="ew")

        for i, label in enumerate(self.doc_labels):
            row = i + 4  # Shift everything down by 1 more row
            tk.Label(self.root, text=label).grid(row=row, column=0, sticky="w", padx=10)

            btn1 = tk.Button(self.root, text="📄", command=lambda l=label: self.load_document(l))
            btn1.grid(row=row, column=1, padx=3, pady=3, sticky="ew")

            btn2 = tk.Button(self.root, text="📥", command=lambda l=label: self.load_dataframe_csv(l))
            btn2.grid(row=row, column=2, padx=3, pady=3, sticky="ew")

            btn3 = tk.Button(self.root, text="💾", command=lambda l=label: self.save_dataframe_csv(l))
            btn3.grid(row=row, column=3, padx=3, pady=3, sticky="ew")

            # Save buttons references if needed
            self.frames[label] = (btn1, btn2, btn3)

            # frame = DragDropFrame(self.root, label, self.store_file, self.export_single_file)
            # frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
            # self.frames[label] = frame

            status_lbl = tk.Label(self.root, text="", font=("Arial", 12))
            status_lbl.grid(row=row, column=4, padx=5)
            self.status_labels[label] = status_lbl

        last_row = len(self.doc_labels) + 4

        tk.Frame(self.root, height=1, bd=1, relief="sunken", bg="gray").grid(
            row=last_row, column=0, columnspan=4, sticky="ew", pady=0
        )
        last_row += 1

        self.generate_btn = tk.Button(
            self.root,
            text="Generar Excel EERR",
            command=self.generate_excel,
            bg="green",
            fg="black",
        )
        self.generate_btn.grid(row=last_row, column=0, columnspan=2, pady=5, padx=(10, 5), sticky="ew")

        self.export_dfs_btn = tk.Button(
            self.root,
            text="💾 Exportar DataFrames",
            command=self.export_all_dataframes,
            bg="blue",
            fg="black",
        )
        self.export_dfs_btn.grid(row=last_row, column=2, columnspan=2, pady=5, padx=(5, 10), sticky="ew")

        # Configure all used columns to expand nicely
        for col in range(4):
            self.root.grid_columnconfigure(col, weight=1)

        button_columns = [1, 2, 3]
        for col in button_columns:
            self.root.grid_columnconfigure(col, weight=1, minsize=120)

    def _add_path_selector(self, label_text, variable, row, file=False):
        tk.Label(self.root, text=label_text).grid(row=row, column=0, sticky="w", padx=10)
        entry = tk.Entry(self.root, textvariable=variable)
        entry.grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 0))
        
        def browse():
            if file:
                path = filedialog.askopenfilename(title=f"Seleccionar {label_text}")
            else:
                path = filedialog.askdirectory(title=f"Seleccionar {label_text}")
            if path:
                variable.set(path)

        browse_btn = tk.Button(self.root, text="📂", command=browse)
        browse_btn.grid(row=row, column=3, padx=5)

    def store_file(self, label, path):
        self.file_inputs[label] = path
        # Cuando se selecciona un archivo, se procesa y guarda el DataFrame
        df = file_to_df(path, label)
        self.data_frames[label] = df

    def export_single_file(self, label, file_path):
        # Intentar usar DataFrame cargado
        df = self.data_frames.get(label)

        if df is None or df.empty:
            # Si no está cargado o vacío, lanzar la GUI para extraer datos
            df = launch_pdf_gui(file_path, label)
            if not df.empty:
                self.data_frames[label] = df  # Guardar para futuros usos

        if df is not None and not df.empty:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title=f"Guardar Excel para {label}"
            )
            if save_path:
                df.to_excel(save_path, index=False)
                messagebox.showinfo("Éxito", f"Archivo guardado como:\n{save_path}")
        else:
            messagebox.showwarning("Vacío", f"No se extrajeron datos para {label}.")

    def generate_excel(self):
        template_path = self.template_path.get()
        if not os.path.isfile(template_path):
            messagebox.showerror("Error", "La plantilla Excel no es válida.")
            return

        wb = load_workbook(template_path)

        for sheet_name, labels in self.sheet_to_labels.items():
            dfs = {label: self.data_frames[label] for label in labels if label in self.data_frames}
            if not dfs:
                continue
            model_path = self.model_dir.get()
            df_to_wb(dfs, wb, sheet_name, model_path)

        #format_EERR(wb)
        output_path = os.path.join(self.output_dir.get(), "EERR.xlsx")
        wb.save(output_path)

        messagebox.showinfo("Éxito", "Reporte Excel generado correctamente!")

        # except Exception as e:
        #     messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    def load_document(self, label):
        if label == "Transbank":
            path = filedialog.askdirectory(title=f"Seleccionar carpeta para {label}")
        else:
            path = filedialog.askopenfilename(title=f"Seleccionar archivo para {label}")
        if path:
            self.file_inputs[label] = path
            df = file_to_df(path, label)
            self.data_frames[label] = df
            self.status_labels[label].config(text="✅", fg="green")

    def load_dataframe_csv(self, label):
        path = filedialog.askopenfilename(title=f"Cargar DataFrame CSV para {label}", filetypes=[("CSV Files", "*.csv")])
        if path:
            try:
                df = pd.read_csv(path)
                self.data_frames[label] = df
                self.status_labels[label].config(text="✅", fg="green")
                #messagebox.showinfo("Cargado", f"DataFrame cargado desde CSV para: {label}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar CSV:\n{e}")

    def save_dataframe_csv(self, label):
        df = self.data_frames.get(label)
        if df is None or df.empty:
            messagebox.showwarning("Vacío", f"No hay DataFrame para: {label}")
            return
        save_path = os.path.join(self.output_dir.get(), f"{label}.csv")
        try:
            df.to_csv(save_path, index=False)
            messagebox.showinfo("Guardado", f"DataFrame guardado como:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar CSV:\n{e}")

    def export_all_dataframes(self):
        if not self.data_frames:
            messagebox.showwarning("Sin datos", "No hay DataFrames para exportar.")
            return

        out_dir = self.output_dir.get()
        os.makedirs(out_dir, exist_ok=True)

        errors = []
        for label, df in self.data_frames.items():
            try:
                if df is not None and not df.empty:
                    df.to_csv(os.path.join(out_dir, f"{label}.csv"), index=False)
            except Exception as e:
                errors.append(f"{label}: {e}")

        if errors:
            messagebox.showerror("Errores", f"Algunos archivos no se pudieron exportar:\n" + "\n".join(errors))
        else:
            messagebox.showinfo("Éxito", f"Todos los DataFrames se guardaron en:\n{out_dir}")


from tkinterdnd2 import TkinterDnD

if __name__ == "__main__":
    root = TkinterDnD.Tk()  # instead of tk.Tk()
    app = ExcelMergerApp(root)
    root.resizable(True, False)
    root.geometry("700x500")
    root.mainloop()
