import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd

class TableExtractorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Selecciona área de tabla PDF")
        self.file_path = None
        self.page_number = 0
        self.start_x = self.start_y = 0
        self.rect = None
        self.image_id = None
        self.nrows = 1
        self.ncols = 1
        self.grid_lines = []
        self.accumulated_df = pd.DataFrame()
        self.col_fractions = [0.0, 1.0]
        self.vertical_correction = 0
        self.debug_labels = []

        # -------- Layout --------
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)

        # Canvas = PDF viewer
        self.canvas = tk.Canvas(self.main_frame, cursor="cross", bg="white")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Right panel
        self.right_panel = tk.Frame(self.main_frame)
        self.right_panel.pack(side="right", fill="y")

        # Buttons
        self.select_button = tk.Button(self.right_panel, text="Cargar PDF", command=self.load_pdf)
        self.select_button.pack()

        self.extract_button = tk.Button(self.right_panel, text="Extraer tabla (Enter)", command=self.extract_grid_table)
        self.extract_button.pack()

        self.save_all_button = tk.Button(self.right_panel, text="Guardar todo", command=self.save_all)
        self.save_all_button.pack()

        self.reset_button = tk.Button(self.right_panel, text="Reiniciar selección", command=self.reset_all)
        self.reset_button.pack()

        # PDF page navigation
        self.nav_frame = tk.Frame(self.right_panel)
        self.nav_frame.pack(pady=5)

        self.prev_button = tk.Button(self.nav_frame, text="⟨ Página", command=self.prev_page)
        self.prev_button.pack(side="left")

        self.page_label = tk.Label(self.nav_frame, text="Página 1")
        self.page_label.pack(side="left")

        self.next_button = tk.Button(self.nav_frame, text="Página ⟩", command=self.next_page)
        self.next_button.pack(side="left")

        # Status label for grid size
        self.status_label = tk.Label(self.right_panel, text="Rows: 1, Cols: 1")
        self.status_label.pack()

        # Treeview table preview
        from tkinter import ttk
        self.preview = ttk.Treeview(self.right_panel, show="headings")
        self.preview.pack(fill="both", expand=True)

        # -------- Bindings --------
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.canvas.bind("<MouseWheel>", self.adjust_column_width)

        self.root.bind("<Up>", self.increase_rows)
        self.root.bind("<Down>", self.decrease_rows)
        self.root.bind("<Right>", self.increase_cols)
        self.root.bind("<Left>", self.decrease_cols)
        self.root.bind("<Return>", self.extract_grid_table)

        self.root.bind("=", self.increase_offset)
        self.root.bind("-", self.decrease_offset)
        self.root.bind("d", self.debug_print_grid_info)
        self.root.bind("p", self.print_cursor_position)

    # def __init__(self, root):
    #     self.root = root
    #     self.root.title("Selecciona área de tabla PDF")
    #     self.file_path = None

    #     self.canvas = tk.Canvas(self.root, cursor="cross")
    #     self.canvas.pack(fill="both", expand=True)

    #     self.start_x = self.start_y = 0
    #     self.rect = None
    #     self.image_id = None
    #     self.page_number = 0
    #     self.nrows = 1
    #     self.ncols = 1
    #     self.grid_lines = []
    #     self.accumulated_df = pd.DataFrame()

    #     self.select_button = tk.Button(self.root, text="Cargar PDF", command=self.load_pdf)
    #     self.select_button.pack()

    #     self.extract_button = tk.Button(self.root, text="Extraer tabla (Enter)", command=self.extract_grid_table)
    #     self.extract_button.pack()

    #     self.save_all_button = tk.Button(self.root, text="Guardar todo", command=self.save_all)
    #     self.save_all_button.pack()

    #     self.reset_button = tk.Button(self.root, text="Reiniciar selección", command=self.reset_all)
    #     self.reset_button.pack()

    #     self.canvas.bind("<Button-1>", self.on_mouse_down)
    #     self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
    #     self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

    #     # Bind keys for grid control
    #     self.root.bind("<Up>", self.increase_rows)
    #     self.root.bind("<Down>", self.decrease_rows)
    #     self.root.bind("<Right>", self.increase_cols)
    #     self.root.bind("<Left>", self.decrease_cols)
    #     self.root.bind("<Return>", self.extract_grid_table)

    #     # Label to show current grid size
    #     self.status_label = tk.Label(self.root, text="Rows: 1, Cols: 1")
    #     self.status_label.pack()

    #     self.col_fractions = [0.0, 1.0]  # initialized with one column
    #     self.canvas.bind("<MouseWheel>", self.adjust_column_width)

    #     self.vertical_correction = 0
    #     self.root.bind("=", self.increase_offset)  # press `=` or `+`
    #     self.root.bind("-", self.decrease_offset)  # press `-`

    #     self.debug_labels = []
    #     self.root.bind("d", self.debug_print_grid_info)

    #     self.root.bind("p", self.print_cursor_position)

    def update_preview(self, df):
        # Clear old preview
        self.preview.delete(*self.preview.get_children())

        # Set new columns
        self.preview["columns"] = list(df.columns)
        self.preview["show"] = "headings"
        for col in df.columns:
            self.preview.heading(col, text=col)

        # Insert rows
        for _, row in df.iterrows():
            self.preview.insert("", "end", values=list(row))

    def load_pdf(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not self.file_path:
            return

        self.doc = fitz.open(self.file_path)
        self.display_page(0)
        # Reset grid & selection when loading new PDF
        self.clear_selection()

    def clear_selection(self):
        if self.rect:
            self.canvas.delete(self.rect)
            self.rect = None
        for line in self.grid_lines:
            self.canvas.delete(line)
        self.grid_lines = []
        self.nrows = 1
        self.ncols = 1
        self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")

    def display_page(self, page_num):
        page = self.doc[page_num]
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas.config(width=img.width, height=img.height)
        if self.image_id:
            self.canvas.delete(self.image_id)
        self.image_id = self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        self.page_height = img.height  # for coordinate conversion
        self.page_width = img.width

    def on_mouse_down(self, event):
        self.start_x, self.start_y = event.x, event.y
        if self.rect:
            self.canvas.delete(self.rect)
            self.rect = None
        for line in self.grid_lines:
            self.canvas.delete(line)
        self.grid_lines = []


        self.nrows = 1  # Reset only rows, keep ncols and col_fractions
        self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")

        # self.nrows, self.ncols = 1, 1  # reset only on new selection start
        # self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red")

    def on_mouse_drag(self, event):
        if self.rect:
            self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)
            self.draw_grid()  # redraw grid on every drag update

    def on_mouse_up(self, event):
        self.end_x, self.end_y = event.x, event.y
        # Do NOT reset rows/cols here
        # Just draw grid once more to finalize
        self.draw_grid()
        self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")

    def draw_grid(self):
        # Remove old grid lines
        for line in self.grid_lines:
            self.canvas.delete(line)
        self.grid_lines = []

        if not self.rect:
            return

        x1, y1, x2, y2 = self.canvas.coords(self.rect)
        left, right = min(x1, x2), max(x1, x2)
        top, bottom = min(y1, y2), max(y1, y2)

        width = right - left
        height = bottom - top

        # Vertical lines (columns)
        for fx in self.col_fractions[1:-1]:
            x = left + fx * width
            line = self.canvas.create_line(x, top, x, bottom, fill='blue', dash=(2, 2))
            self.grid_lines.append(line)

        # Horizontal lines (rows)
        for j in range(1, self.nrows):
            y = top + j * height / self.nrows
            line = self.canvas.create_line(left, y, right, y, fill='blue', dash=(2, 2))
            self.grid_lines.append(line)

    def increase_rows(self, event=None):
        if not self.rect:
            return
        if self.nrows < 50:
            self.nrows += 1
            self.draw_grid()
            self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")

    def decrease_rows(self, event=None):
        if not self.rect:
            return
        if self.nrows > 1:
            self.nrows -= 1
            self.draw_grid()
            self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")

    def increase_cols(self, event=None):
        if not self.rect:
            return
        if self.ncols < 50:
            self.ncols += 1
            self.col_fractions = [i / self.ncols for i in range(self.ncols + 1)]
            self.draw_grid()
            self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")

    def decrease_cols(self, event=None):
        if not self.rect:
            return
        if self.ncols > 1:
            self.ncols -= 1
            self.col_fractions = [i / self.ncols for i in range(self.ncols + 1)]
            self.draw_grid()
            self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")

    def adjust_column_width(self, event):
        if not self.rect or self.ncols < 2:
            return

        x1, _, x2, _ = self.canvas.coords(self.rect)
        left, right = min(x1, x2), max(x1, x2)
        width = right - left

        # Relative mouse x position
        x_mouse = event.x
        rel_x = (x_mouse - left) / width

        # Determine which column the mouse is inside
        for i in range(self.ncols):
            fx1 = self.col_fractions[i]
            fx2 = self.col_fractions[i + 1]
            if fx1 <= rel_x < fx2:
                delta = 0.01 if event.delta > 0 else -0.01
                new_fx = self.col_fractions[i + 1] + delta

                # Clamp so it doesn’t cross over the next boundary
                if self.col_fractions[i + 1] + delta < self.col_fractions[i] + 0.005:
                    return
                if i + 2 < len(self.col_fractions) and new_fx > self.col_fractions[i + 2] - 0.005:
                    return

                self.col_fractions[i + 1] = new_fx
                self.draw_grid()
                break

    def increase_offset(self, event=None):
        self.vertical_correction += 1
        print(f"Vertical correction: {self.vertical_correction}px")
        self.draw_grid()

    def decrease_offset(self, event=None):
        self.vertical_correction -= 1
        print(f"Vertical correction: {self.vertical_correction}px")
        self.draw_grid()

    def print_cursor_position(self, event=None):
        # Get mouse position relative to canvas
        x_canvas = self.canvas.winfo_pointerx() - self.canvas.winfo_rootx()
        y_canvas = self.canvas.winfo_pointery() - self.canvas.winfo_rooty()

        if not hasattr(self, 'doc') or not self.doc:
            print("PDF not loaded.")
            return

        try:
            # Get PDF page using PyMuPDF (fitz)
            page = self.doc[self.page_number]
            pix = page.get_pixmap()
            img_width, img_height = pix.width, pix.height

            # Get PDF size using pdfplumber
            with pdfplumber.open(self.file_path) as pdf:
                pdf_page = pdf.pages[self.page_number]
                pdf_width, pdf_height = pdf_page.width, pdf_page.height

            # Scale conversion from Canvas/Image → PDF coords
            scale_x = pdf_width / img_width
            scale_y = pdf_height / img_height

            x_pdf = x_canvas * scale_x
            y_pdf = (img_height - y_canvas) * scale_y  # Y-axis flipped

            print(f"\n🖱 Cursor position:")
            print(f"   📐 Canvas: ({x_canvas:.2f}, {y_canvas:.2f})")
            print(f"   📄 PDF   : ({x_pdf:.2f}, {y_pdf:.2f})\n")

        except Exception as e:
            print(f"[ERROR] Could not get cursor PDF position: {e}")

    def debug_print_grid_info(self, event=None):
        if not self.rect:
            print("No hay selección activa.")
            return

        # Borrar etiquetas anteriores
        for label in self.debug_labels:
            self.canvas.delete(label)
        self.debug_labels = []

        x1, y1, x2, y2 = self.canvas.coords(self.rect)
        top, bottom = min(y1, y2), max(y1, y2)
        left, right = min(x1, x2), max(x1, x2)

        width = right - left
        height = bottom - top

        try:
            with pdfplumber.open(self.file_path) as pdf:
                page = pdf.pages[self.page_number]
                page_height = page.height

                for row in range(self.nrows):
                    cell_top = top + (self.nrows - row - 1) * height / self.nrows
                    cell_bottom = top + (self.nrows - row) * height / self.nrows
                    for col in range(self.ncols):
                        fx1 = self.col_fractions[col]
                        fx2 = self.col_fractions[col + 1]
                        cell_left = left + fx1 * width
                        cell_right = left + fx2 * width

                        bbox_pdf = (
                            cell_left,
                            page_height - cell_bottom,
                            cell_right,
                            page_height - cell_top
                        )
                        print(f"Row {row}, Col {col} -> Canvas bbox: "
                            f"({cell_left:.2f}, {cell_top:.2f}, {cell_right:.2f}, {cell_bottom:.2f}), "
                            f"PDF bbox: {bbox_pdf}")

                        # Mostrar etiquetas visuales en el canvas
                        center_x = (cell_left + cell_right) / 2
                        center_y = (cell_top + cell_bottom) / 2
                        label = self.canvas.create_text(center_x, center_y, text=f"{row},{col}", fill="red", font=("Arial", 8, "bold"))
                        self.debug_labels.append(label)
        except Exception as e:
            print(f"[DEBUG] Error: {e}")

    def extract_grid_table(self, event=None):
        if not self.rect:
            messagebox.showwarning("Área no seleccionada", "Debes seleccionar un área sobre el PDF.")
            return
        
        page = self.doc[self.page_number]
        pix = page.get_pixmap()
        img_width, img_height = pix.width, pix.height

        x1, y1, x2, y2 = self.canvas.coords(self.rect)
        top, bottom = min(y1, y2), max(y1, y2)
        left, right = min(x1, x2), max(x1, x2)

        width = right - left
        height = bottom - top

        if width < 5 or height < 5:
            messagebox.showwarning("Área muy pequeña", "El área seleccionada es demasiado pequeña.")
            return

        try:
            with pdfplumber.open(self.file_path) as pdf:
                page = pdf.pages[self.page_number]
                canvas_width, canvas_height = self.page_width, self.page_height  # from display_page
                pdf_width, pdf_height = page.width, page.height

                scale_x = pdf_width / canvas_width
                scale_y = pdf_height / canvas_height

                # print(f'Canvas (w, h) = ({canvas_width}, {canvas_height})')
                # print(f'PDF (w, h) = ({pdf_width}, {pdf_height})')

                data = []
                for row in range(self.nrows):
                    row_data = []

                    canvas_cell_top = top + row * height / self.nrows
                    canvas_cell_bottom = top + (row + 1) * height / self.nrows


                    for col in range(self.ncols):
                        fx1 = self.col_fractions[col]
                        fx2 = self.col_fractions[col + 1]
                        canvas_cell_left = left + fx1 * width
                        canvas_cell_right = left + fx2 * width

                        # Flip Y and scale everything to PDF space
                        pdf_cell_left = canvas_cell_left * scale_x
                        pdf_cell_right = canvas_cell_right * scale_x

                        bbox = (
                            canvas_cell_left,
                            canvas_cell_top,
                            canvas_cell_right,
                            canvas_cell_bottom
                        )
                        
                        print(f"Row {row}, Col {col} -> Canvas bbox: "
                            f"({canvas_cell_left:.2f}, {canvas_cell_top:.2f}, {canvas_cell_right:.2f}, {canvas_cell_bottom:.2f}), "
                            f"PDF bbox: {bbox}")
                        cropped = page.within_bbox(bbox)
                        texts = cropped.extract_text()
                        row_data.append(texts.strip() if texts else "")
                    data.append(row_data)

            df = pd.DataFrame(data)
            print(df.to_string())
            if not self.accumulated_df.empty:
                if df.shape[1] != self.accumulated_df.shape[1]:
                    messagebox.showwarning("Columnas incompatibles", "El número de columnas no coincide con selecciones anteriores.")
                    return
                self.accumulated_df = pd.concat([self.accumulated_df, df], ignore_index=True)
            else:
                self.accumulated_df = df
            self.update_preview(self.accumulated_df)
            messagebox.showinfo("Tabla añadida", "Se ha añadido esta selección a la tabla acumulada.")

            # df = pd.DataFrame(data)
            # print(df.to_string())
            df.to_excel("tabla_grid_extraida.xlsx", index=False)
            messagebox.showinfo("Éxito", "Tabla extraída y guardada como 'tabla_grid_extraida.xlsx'.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo extraer la tabla:\n{e}")

    def save_all(self):
        if self.accumulated_df.empty:
            messagebox.showwarning("Nada para guardar", "No hay datos acumulados para guardar.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            self.accumulated_df.to_excel(save_path, index=False)
            messagebox.showinfo("Guardado", f"Tabla guardada como {save_path}")

    def reset_all(self):
        self.accumulated_df = pd.DataFrame()
        self.clear_selection()
        messagebox.showinfo("Reiniciado", "Selecciones anteriores eliminadas.")

    def prev_page(self):
        if self.page_number > 0:
            self.page_number -= 1
            self.display_page(self.page_number)
            self.page_label.config(text=f"Página {self.page_number + 1}")

    def next_page(self):
        if self.page_number < len(self.doc) - 1:
            self.page_number += 1
            self.display_page(self.page_number)
            self.page_label.config(text=f"Página {self.page_number + 1}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TableExtractorApp(root)
    root.mainloop()
