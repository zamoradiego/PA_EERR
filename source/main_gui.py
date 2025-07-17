import tkinter as tk
from tkinter import filedialog, messagebox
import os
from openpyxl import load_workbook
from processors import df_to_wb, file_to_df, launch_pdf_gui  # Asumiendo que estos están implementados
from tkinterdnd2 import DND_FILES, TkinterDnD

class DragDropFrame(tk.Frame):
    def __init__(self, parent, label, callback, export_callback):
        super().__init__(parent, bd=2, relief="groove", height=40)
        self.label = label
        self.callback = callback
        self.export_callback = export_callback
        self.file_path = None

        self.grid_columnconfigure(0, weight=1)

        self.label_widget = tk.Label(self, text="Drop or click to select", anchor="w")
        self.label_widget.grid(row=0, column=0, sticky="ew", padx=5)

        export_btn = tk.Button(self, text="Exportar Excel", command=self.export_file)
        export_btn.grid(row=0, column=1, padx=5)

        self.label_widget.bind("<Button-1>", self.on_click)
        self.bind("<Button-1>", self.on_click)

        # Add drag and drop binding
        self.label_widget.drop_target_register(DND_FILES)
        self.label_widget.dnd_bind('<<Drop>>', self.drop)

    def drop(self, event):
        files = self.winfo_toplevel().tk.splitlist(event.data)
        if files:
            file_path = files[0]
            self.set_file(file_path)

    def on_click(self, event):
        file_path = filedialog.askopenfilename(
            title=f"Seleccionar archivo para {self.label}",
            filetypes=[("Archivos válidos", ("*.xlsx", "*.xls", "*.csv", "*.pdf")), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.set_file(file_path)

    def set_file(self, file_path):
        self.file_path = file_path
        self.label_widget.config(text=os.path.basename(file_path))
        self.callback(self.label, file_path)

    def export_file(self):
        if not self.file_path:
            messagebox.showerror("Error", f"No hay archivo cargado para {self.label}")
            return
        self.export_callback(self.label, self.file_path)


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

        # Crear los frames con botón exportar individual
        for i, label in enumerate(self.doc_labels):
            tk.Label(self.root, text=label).grid(row=i, column=0, sticky="w", padx=10)
            frame = DragDropFrame(self.root, label, self.store_file, self.export_single_file)
            frame.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
            self.frames[label] = frame

        self.generate_btn = tk.Button(
            self.root,
            text="Generar Excel Consolidado",
            command=self.generate_excel,
            bg="green",
            fg="white",
        )
        self.generate_btn.grid(row=len(self.doc_labels) + 1, column=0, columnspan=2, pady=20, sticky="ew")

        # Configuramos para que las columnas 0 y 1 se expandan bien
        self.root.grid_columnconfigure(0, weight=0)  # etiqueta
        self.root.grid_columnconfigure(1, weight=1)  # dragdrop + botón

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
        template_path = os.path.join('.', 'template_EERR.xlsx')
        wb = load_workbook(template_path)

        for sheet_name, labels in self.sheet_to_labels.items():
            dfs = {label: self.data_frames[label] for label in labels if label in self.data_frames}
            if not dfs:
                continue
            df_to_wb(dfs, wb, sheet_name)

        #format_EERR(wb)
        os.makedirs("outputs", exist_ok=True)
        output_path = os.path.join("outputs", "EERR.xlsx")
        wb.save(output_path)

        messagebox.showinfo("Éxito", "Reporte Excel generado correctamente!")

        # except Exception as e:
        #     messagebox.showerror("Error", f"Ocurrió un error:\n{e}")


from tkinterdnd2 import TkinterDnD

if __name__ == "__main__":
    root = TkinterDnD.Tk()  # instead of tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
