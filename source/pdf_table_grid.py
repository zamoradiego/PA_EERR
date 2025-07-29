import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd

class TableExtractorApp:

    def __init__(self, root, file_path=None):
        self.root = root
        self.root.title("Selecciona área de tabla PDF")
        self.file_path = file_path
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
        self.result_df = None

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

        self.confirm_button = tk.Button(self.right_panel, text="Confirmar selección", command=self.confirm_selection)
        self.confirm_button.pack()

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

        # # Status label for grid size
        self.status_label = tk.Label(self.right_panel, text="Rows: 1, Cols: 1")
        self.status_label.pack()

        status_frame = tk.Frame(self.root)
        status_frame.pack()

        self.rows_var = tk.StringVar(value="1")
        self.cols_var = tk.StringVar(value="1")

        tk.Label(status_frame, text="Rows:").pack(side="left")
        self.rows_entry = tk.Entry(status_frame, width=4, textvariable=self.rows_var)
        self.rows_entry.pack(side="left")

        tk.Label(status_frame, text="Cols:").pack(side="left")
        self.cols_entry = tk.Entry(status_frame, width=4, textvariable=self.cols_var)
        self.cols_entry.pack(side="left")

        # React to typing immediately
        self.rows_var.trace_add("write", self.update_grid_dimensions)
        self.cols_var.trace_add("write", self.update_grid_dimensions)

        # Treeview table preview
        from tkinter import ttk
        self.preview = ttk.Treeview(self.right_panel, show="headings")
        self.preview.pack(fill="both", expand=True)
        self.preview.bind("<Double-Button-1>", self.edit_header)

        self.header_mode = False
        self.header_col_fractions = []  # will store header grid
        self.header_text = None
        self.root.bind("h", self.activate_header_mode)

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

        self.main_frame.bind("<Configure>", self.resize_right_panel)

        if self.file_path:
            self.doc = fitz.open(self.file_path)
            self.display_page(0)

    def resize_right_panel(self, event=None):
        try:
            total_width = self.main_frame.winfo_width()
            right_width = int(total_width / 2.5)  # 1.5x smaller than left
            self.right_panel.config(width=right_width)
        except:
            pass

    def update_grid_dimensions(self, *args):
        try:
            nrows = int(self.rows_var.get())
            ncols = int(self.cols_var.get())

            if nrows > 0 and ncols > 0:
                self.nrows = nrows
                self.ncols = ncols
                self.update_col_fractions()
                self.draw_grid()
        except ValueError:
            # User is mid-typing, e.g., "1a"
            pass

    def update_col_fractions(self):
        # Recalculate column fractions (evenly spaced)
        self.col_fractions = [i / self.ncols for i in range(self.ncols + 1)]


    def activate_header_mode(self, event=None):
        self.header_mode = True
        # Preserve column structure before clearing visuals
        self.header_col_fractions = self.col_fractions.copy()
        ncols = self.ncols  # remember the number of columns
        self.clear_selection()

        # Restore previous column config
        self.col_fractions = self.header_col_fractions.copy()
        self.ncols = ncols

        self.status_label.config(text="Selecciona el área del encabezado (1 fila, columnas ajustables)")

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
        # Do not reset ncols or col_fractions here
        self.nrows = 1  # Only reset rows
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

        self.status_label.config(text=f"Rows: {self.nrows}, Cols: {self.ncols}")
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
            line = self.canvas.create_line(x, top, x, bottom, fill='red')
            self.grid_lines.append(line)

        # Horizontal lines (rows)
        for j in range(1, self.nrows):
            y = top + j * height / self.nrows
            line = self.canvas.create_line(left, y, right, y, fill='red')
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

    def extract_grid_table(self, event=None):

        if not self.rect:
            messagebox.showwarning("Área no seleccionada", "Debes seleccionar un área sobre el PDF.")
            return

        else:
            with pdfplumber.open(self.file_path) as pdf:
                page = pdf.pages[self.page_number]
                canvas_width, canvas_height = self.page_width, self.page_height  # from display_page
                pdf_width, pdf_height = page.width, page.height

                scale_x = pdf_width / canvas_width
                scale_y = pdf_height / canvas_height

                x1, y1, x2, y2 = self.canvas.coords(self.rect)
                top, bottom = min(y1, y2), max(y1, y2)
                left, right = min(x1, x2), max(x1, x2)

                width = right - left
                height = bottom - top

                if width < 5 or height < 5:
                    messagebox.showwarning("Área muy pequeña", "El área seleccionada es demasiado pequeña.")
                    return
            
                if self.header_mode:
                    self.header_mode = False  # Reset mode after selection
                    self.header_col_fractions = self.col_fractions.copy()
                    header_data = []

                    canvas_cell_top = top
                    canvas_cell_bottom = top + height / self.nrows

                    for col in range(self.ncols):
                        fx1 = self.col_fractions[col]
                        fx2 = self.col_fractions[col + 1]
                        canvas_cell_left = left + fx1 * width
                        canvas_cell_right = left + fx2 * width
                        padding_y = 3
                        padding_x = 3

                        bbox = (
                            canvas_cell_left - padding_x,
                            canvas_cell_top - padding_y,
                            canvas_cell_right + padding_x,
                            canvas_cell_bottom + padding_y
                        )

                        cropped = page.within_bbox(bbox)
                        text = cropped.extract_text()
                        text = text.replace('\n', ' ')
                        #print(text)
                        header_data.append(text.strip() if text else f"Col{col+1}")

                    self.header_text = header_data
                    self.show_info("Encabezado definido", f"Encabezado seleccionado: {self.header_text}")

                    # 🔄 If there's already data, update column headers in the accumulated DataFrame and preview
                    if not self.accumulated_df.empty:
                        if len(self.header_text) == self.accumulated_df.shape[1]:
                            self.accumulated_df.columns = self.header_text
                            self.update_preview(self.accumulated_df)
                        else:
                            messagebox.showwarning("Columnas incompatibles", "El encabezado no coincide en número de columnas con la tabla actual.")
                    return
                
                else:
                    try:
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

                                padding_y = 3
                                padding_x = 4

                                bbox = (
                                    canvas_cell_left,
                                    canvas_cell_top - padding_y,
                                    canvas_cell_right,
                                    canvas_cell_bottom + padding_y
                                )
                                
                                #print(f"Row {row}, Col {col} -> Canvas bbox: "
                                #    f"({canvas_cell_left:.2f}, {canvas_cell_top:.2f}, {canvas_cell_right:.2f}, {canvas_cell_bottom:.2f}), "
                                #    f"PDF bbox: {bbox}")
                                cropped = page.within_bbox(bbox)
                                texts = cropped.extract_text()
                                row_data.append(texts.strip() if texts else "")
                            data.append(row_data)

                        df = pd.DataFrame(data, columns=self.header_text if self.header_text else None)
                        #print(df.to_string())
                        if not self.accumulated_df.empty:
                            if df.shape[1] != self.accumulated_df.shape[1]:
                                messagebox.showwarning("Columnas incompatibles", "El número de columnas no coincide con selecciones anteriores.")
                                return
                            self.accumulated_df = pd.concat([self.accumulated_df, df], ignore_index=True)
                        else:
                            self.accumulated_df = df
                        self.update_preview(self.accumulated_df)
                        self.show_info("Tabla añadida", "Se ha añadido esta selección a la tabla acumulada.")

                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo extraer la tabla:\n{e}")

        
    def update_grid_size(self, event=None):
        try:
            rows = int(self.rows_entry.get())
            cols = int(self.cols_entry.get())
            if rows > 0:
                self.nrows = rows
            if cols > 0:
                self.ncols = cols
                self.col_fractions = [i / cols for i in range(cols + 1)]
            self.draw_grid()
        except ValueError:
            messagebox.showwarning("Entrada inválida", "Debes ingresar números válidos.")

    # def save_all(self):
    #     if self.accumulated_df.empty:
    #         messagebox.showwarning("Nada para guardar", "No hay datos acumulados para guardar.")
    #         return

    #     save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
    #                                             filetypes=[("Excel files", "*.xlsx")])
    #     if save_path:
    #         self.accumulated_df.to_excel(save_path, index=False)
    #         self.show_info("Guardado", f"Tabla guardada como {save_path}")

    def confirm_selection(self):
        df = self.accumulated_df

        if self.accumulated_df.empty:
            messagebox.showerror("Error", "No se extrajo ningún dato.")
        else:
            self.result_df = df
            self.root.destroy()

    def reset_all(self):
        self.accumulated_df = pd.DataFrame()
        self.update_preview(self.accumulated_df)
        self.clear_selection()
        self.show_info("Reiniciado", "Selecciones anteriores eliminadas.")

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

    def show_info(self, title, message):
        messagebox.showinfo(title, message)
        self.canvas.focus_set()
        self.root.focus_force()

    def edit_header(self, event):
        # Identify which column header was clicked
        region = self.preview.identify_region(event.x, event.y)
        if region != "heading":
            return

        column_id = self.preview.identify_column(event.x)
        col_index = int(column_id.replace("#", "")) - 1

        old_name = self.preview["columns"][col_index]

        # Get x/y coordinates of the header cell
        bbox = self.preview.bbox("")

        if not bbox:
            return

        x, y, width, height = self.preview.bbox(column_id)

        # Create Entry widget for editing
        entry = tk.Entry(self.preview)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, old_name)
        entry.focus()

        def save_header_change(event=None):
            new_name = entry.get().strip()
            entry.destroy()
            if not new_name:
                return
            # Update Treeview columns
            columns = list(self.preview["columns"])
            columns[col_index] = new_name
            self.preview["columns"] = columns
            self.preview.heading(f"#{col_index+1}", text=new_name)

            # Update accumulated_df column names
            if not self.accumulated_df.empty:
                df_columns = list(self.accumulated_df.columns)
                df_columns[col_index] = new_name
                self.accumulated_df.columns = df_columns

            # Update stored header text for future extractions
            if self.header_text and len(self.header_text) == len(columns):
                self.header_text[col_index] = new_name

        entry.bind("<Return>", save_header_change)
        entry.bind("<FocusOut>", lambda e: entry.destroy())


if __name__ == "__main__":
    root = tk.Tk()
    app = TableExtractorApp(root)
    root.mainloop()
