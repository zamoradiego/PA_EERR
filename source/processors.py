# processors.py
import pandas as pd
import pickle
from text_classification import classify_text
from pdf_table_grid import TableExtractorApp
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1
import re
import os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import column_index_from_string
from openpyxl.utils import range_boundaries, get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Alignment, Protection
from datetime import datetime
from openpyxl.worksheet.datavalidation import DataValidation
from os.path import join
import numpy as np

PATH_MODEL = join('..', 'prediccion')

columns_to_extract = {'Cuenta BCI 18' : ['Fecha', 'Sucursal', 'Descripción', 'N° Documento', 'Cheques y otros cargos' , 'Depósitos y Abono', 'Saldo diario'],
                      'Cuenta BCI 85' : ['Fecha', 'Sucursal', 'Descripción', 'N° Documento', 'Cheques y otros cargos' , 'Depósitos y Abono', 'Saldo diario'],
                      'Siteminder' : ['Referencia de la reserva', 'Nombres de los huéspedes', 'Llegada', 'Salida', 'Canal', 'Affiliated Channel', 'Referral', 'Habitación', 'Fecha de reserva', 'Estado de la reserva',	'Ocupación', 'Precio total'],
                      'Tarjeta de crédito internacional' : ['NUMERO REFERENCIA INTERNACIONAL', 'FECHA OPERACIÓN', 'DESCRIPCION OPERACION O COBRO', 'CIUDAD', 'PAIS', 'MONTO MONEDA ORIGEN', 'MONTO US$'],
                      'Tarjeta de crédito nacional 24' : ['LUGAR DE OPERACIÓN', 'FECHA OPERACIÓN', 'CODIGO REFERENCIA', 'DESCRIPCION OPERACION O COBRO', 'MONTO OPERACIÓN O COBRO', 'MONTO TOTAL A PAGAR'],
                      'Tarjeta de crédito nacional 69' : ['LUGAR DE OPERACIÓN', 'FECHA OPERACIÓN', 'CODIGO REFERENCIA', 'DESCRIPCION OPERACION O COBRO', 'MONTO OPERACIÓN O COBRO', 'MONTO TOTAL A PAGAR'],
                      'Cuenta BCI Comercio Exterior' : ['FECHA', 'SUCURSAL', 'DESCRIPCION', 'N° DE DOCUMENTO', 'CHEQUES Y OTROS CARGOS', 'DEPOSITOS Y ABONOS', 'SALDO DIARIO'],
                      'Banco Security' : ['fecha ', 'descripción ', 'número de documentos ', 'cargos ', 'abonos ', 'saldos '],
                      'Transbank' : ['Fecha de venta - nan', 'Cód. comercio - nan',
                                      'Nombre local - nan', 'Tipo de movimiento - nan',
                                      'Tipo de tarjeta - nan', 'Identificador - N° de tarjeta',
                                      'nan - Cód. de autorización', 'nan - Orden de pedido',
                                      'nan - Número único', 'nan - ID de servicio', 'Tipo de cuota - nan',
                                      'Monto afecto - nan', 'Monto exento - nan', 'N° de cuotas - nan',
                                      'Monto cuota - nan', 'Fecha de abono - nan', 'N° de boleta - nan',
                                      'Monto vuelto - nan'],
                      'Transferencias' : ['Monto $', 'Nombre Destino/Origen', 'Mensaje Destino']}

column_dictionary = {'Banco Security' : {'fecha ' : 'FECHA', 
                                   'descripción ' : 'DESCRIPCION', 
                                   'número de documentos ' : 'Nº DOCUMENTO', 
                                   'cargos ' : 'CARGO', 
                                   'abonos ' : 'ABONO', 
                                   'saldos ' : 'SALDO'},
                     'Siteminder' : {'Referencia de la reserva' : 'ID RESERVA', 
                                     'Nombres de los huéspedes' : 'HUESPEDES', 
                                     'Llegada' : 'LLEGADA', 
                                     'Salida' : 'SALIDA',
                                     'Canal' : 'CANAL', 
                                     'Affiliated Channel' :'CANAL AFILIADO', 
                                     'Estado de la reserva' : 'ESTADO', 
                                     #'Ocupación' : 'Habitaciones', 
                                     'Habitación' : 'HABITACIONES', 
                                     'Precio total' : 'TOTAL',
                                     'Fecha de reserva' : 'FECHA RESERVA'},
                     'Transbank' : {'Fecha de venta - nan' : 'FECHA', 
                                    'Cód. comercio - nan' : 'CODIGO COMERCIO',
                                    'Nombre local - nan' : 'LOCAL',
                                    'Tipo de movimiento - nan' : 'TIPO MOVIMIENTO',
                                    'Tipo de tarjeta - nan': 'TARJETA',
                                    'Identificador - N° de tarjeta' : 'NUMERO TARJETA',
                                    'nan - Cód. de autorización' : 'CODIGO AUTORIZACION',
                                    'nan - Orden de pedido' : 'ORDEN',
                                    'nan - Número único' : 'NUMERO OPERACION',
                                    'nan - ID de servicio' : 'ID SERVICIO',
                                    'Tipo de cuota - nan' : 'TIPO CUOTA',
                                    'Monto afecto - nan' : 'TOTAL',
                                    'Monto exento - nan' : 'MONTO',
                                    'N° de cuotas - nan' : 'NUMERO CUOTAS',
                                    'Monto cuota - nan' : 'MONTO CUOTA',
                                    'Fecha de abono - nan' : 'FECHA ABONO',
                                    'N° de boleta - nan' : 'NUMERO BOLETA',
                                    'Monto vuelto - nan' : 'VUELTO'},
                     'Transferencias' : {'Monto $' : 'CARGO TRANSF', 
                                         'Nombre Destino/Origen' : 'NOMBRE DESTINO', 
                                         'Mensaje Destino' : 'MENSAJE'},
                     'Cuenta BCI Comercio Exterior' : {'SUCURSAL' : 'OFICINA', 
                                                       'CHEQUES Y OTROS CARGOS' : 'CARGO', 
                                                       'DEPOSITOS Y ABONOS' : 'ABONO', 
                                                       'SALDO DIARIO' : 'SALDO'},
                     'Cuenta BCI 85' : {'Fecha' : 'FECHA', 
                                        'Sucursal' : 'OFICINA', 
                                        'Descripción' : 'MOVIMIENTO', 
                                        'N° Documento': 'N° DOCUMENTO', 
                                        'Cheques y otros cargos' : 'CARGO',
                                        'Depósitos y Abono' : 'ABONO',
                                        'Saldo diario' : 'SALDO'},
                     'Cuenta BCI 18' : {'Fecha' : 'FECHA', 
                                        'Sucursal' : 'OFICINA', 
                                        'Descripción' : 'DESCRIPCION', 
                                        'N° Documento': 'Nº DOCUMENTO', 
                                        'Cheques y otros cargos' : 'CARGO',
                                        'Depósitos y Abono' : 'ABONO',
                                        'Saldo diario' : 'SALDO'},
                     'Tarjeta de crédito nacional 24' : {'FECHA OPERACIÓN' : 'FECHA', 
                                                         'CODIGO REFERENCIA' : 'CÓDIGO REFERENCIA', 
                                                         'LUGAR DE OPERACIÓN' : 'LUGAR OPERACIÓN', 
                                                         'DESCRIPCION OPERACION O COBRO': 'DESCRIPCION', 
                                                         'MONTO OPERACIÓN O COBRO' : 'CARGO'},
                     'Tarjeta de crédito nacional 69' : {'FECHA OPERACIÓN' : 'FECHA', 
                                                         'CODIGO REFERENCIA' : 'CÓDIGO REFERENCIA', 
                                                         'LUGAR DE OPERACIÓN' : 'LUGAR OPERACIÓN', 
                                                         'DESCRIPCION OPERACION O COBRO': 'DESCRIPCION', 
                                                         'MONTO OPERACIÓN O COBRO' : 'CARGO'},
                     'Tarjeta de crédito internacional' : {'FECHA OPERACIÓN' : 'FECHA', 
                                                           'NUMERO REFERENCIA INTERNACIONAL' : 'CÓDIGO REFERENCIA', 
                                                           'DESCRIPCION OPERACION O COBRO' : 'DESCRIPCION'}}

def convert_date(row):
    #return datetime.strptime(row, "%m/%d/%Y")
    return datetime.strptime(row, "%Y-%m-%d")


def move_named_range(ws, wb, named_range_name, offset_rows):
    """
    Mueve un rango nombrado verticalmente en la hoja, preservando las fórmulas
    contenidas dentro del rango. También actualiza la definición del named range.

    Parámetros:
    - ws: worksheet de openpyxl donde está el rango
    - wb: workbook de openpyxl para actualizar el named range
    - named_range_name: str, nombre del rango definido (ej. 'Resumen_BCI')
    - offset_rows: int, número de filas a desplazar (positivo hacia abajo, negativo hacia arriba)
    """
    # Obtener el named range definido
    defined_name = wb.defined_names.get(named_range_name)
    if defined_name is None:
        raise ValueError(f"Named range '{named_range_name}' no encontrado")

    # Obtener la referencia actual
    attr = defined_name.attr_text
    sheet_name, cell_range = attr.split('!')
    sheet_name = sheet_name.strip("'")
    cell_range = cell_range.replace('$', '')

    if sheet_name != ws.title:
        raise ValueError(f"Named range '{named_range_name}' no está en la hoja '{ws.title}'")

    # Parsear coordenadas del rango
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)

    # Guardar fórmulas originales dentro del rango
    formula_map = {}
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            if cell.data_type == 'f' or (isinstance(cell.value, str) and cell.value.startswith('=')):
                formula_map[(row + offset_rows, col)] = cell.value  # guardar con nueva posición

    # Construir el rango actual
    full_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"

    # Mover físicamente el contenido del rango
    ws.move_range(full_range, rows=offset_rows, cols=0)

    # Reinsertar las fórmulas en su nueva posición
    for (new_row, col), formula in formula_map.items():
        ws.cell(row=new_row, column=col, value=formula)

    # Actualizar la definición del named range
    new_min_row = min_row + offset_rows
    new_max_row = max_row + offset_rows
    new_range = f"'{ws.title}'!${get_column_letter(min_col)}${new_min_row}:${get_column_letter(max_col)}${new_max_row}"
    defined_name.attr_text = new_range

def reapply_data_validations(ws, table, start_row, num_new_rows):
    """Clone data validations from the original row and apply to the rest of the table."""
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    data_validations = list(ws.data_validations.dataValidation)

    for dv in data_validations:
        # For each cell in the original row where validation is applied
        for cell_range in dv.ranges:
            for col in range(min_col, max_col + 1):
                cell = ws.cell(row=min_row + 1, column=col)  # +1 to skip headers
                col_letter = get_column_letter(col)
                ref_cell = f"{col_letter}{min_row + 1}"
                if ref_cell in cell_range:
                    # Reapply validation to this entire column from start_row to end
                    new_range = f"{col_letter}{start_row + 1}:{col_letter}{start_row + num_new_rows}"
                    new_dv = DataValidation(
                        type=dv.type,
                        formula1=dv.formula1,
                        formula2=dv.formula2,
                        showDropDown=dv.showDropDown,
                        showInputMessage=dv.showInputMessage,
                        showErrorMessage=dv.showErrorMessage,
                        promptTitle=dv.promptTitle,
                        prompt=dv.prompt,
                        errorTitle=dv.errorTitle,
                        error=dv.error,
                        operator=dv.operator,
                        allowBlank=dv.allowBlank
                    )
                    new_dv.add(new_range)
                    ws.add_data_validation(new_dv)

# def write_df_to_named_table_by_header(ws, sheet_name, table_name, df):
#     table = ws.tables.get(table_name)
#     if table is None:
#         raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'.")

#     start_cell, end_cell = table.ref.split(':')
#     start_col_letter = ''.join(filter(str.isalpha, start_cell))
#     start_col = column_index_from_string(start_col_letter)
#     start_row = int(''.join(filter(str.isdigit, start_cell)))

#     headers = [ws.cell(row=start_row, column=col_idx).value for col_idx in range(start_col, start_col + len(table.tableColumns))]
#     header_to_col = {header: col_idx for header, col_idx in zip(headers, range(start_col, start_col + len(headers)))}

#     for row_offset, (_, df_row) in enumerate(df.iterrows(), start=1):

#         for header, col_idx in header_to_col.items():
#             cell = ws.cell(row=start_row + 1, column=col_idx)  # check the formula in the template row
#             if header in df.columns and not cell.data_type == 'f':
#                 ws.cell(row=start_row + row_offset, column=col_idx, value=df_row[header])

#     new_end_row = start_row + len(df)
#     new_ref = f"{start_cell}:{ws.cell(row=new_end_row, column=start_col + len(headers) - 1).coordinate}"
#     table.ref = new_ref
#     reapply_data_validations(ws, table, start_row, len(df))

#     return new_end_row

def write_df_to_named_table_by_header(ws, sheet_name, table_name, df):
    table = ws.tables.get(table_name)
    if table is None:
        raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'.")

    start_cell, end_cell = table.ref.split(':')
    start_col_letter = ''.join(filter(str.isalpha, start_cell))
    start_col = column_index_from_string(start_col_letter)
    start_row = int(''.join(filter(str.isdigit, start_cell)))

    headers = [ws.cell(row=start_row, column=col_idx).value for col_idx in range(start_col, start_col + len(table.tableColumns))]
    header_to_col = {header: col_idx for header, col_idx in zip(headers, range(start_col, start_col + len(headers)))}

    for row_offset, (_, df_row) in enumerate(df.iterrows(), start=1):
        for header, col_idx in header_to_col.items():
            cell = ws.cell(row=start_row + 1, column=col_idx)  # check the formula in the template row
            if header in df.columns and not cell.data_type == 'f':
                ws.cell(row=start_row + row_offset, column=col_idx, value=df_row[header])

    new_end_row = start_row + len(df)
    new_ref = f"{start_cell}:{ws.cell(row=new_end_row, column=start_col + len(headers) - 1).coordinate}"
    table.ref = new_ref
    reapply_data_validations(ws, table, start_row, len(df))

    # --- COPY FORMULAS ---
    first_data_row = start_row + 1
    for col_idx in range(start_col, start_col + len(headers)):
        source_cell = ws.cell(row=first_data_row, column=col_idx)
        if source_cell.data_type == 'f':  # formula
            for row_offset in range(1, len(df)):
                target_cell = ws.cell(row=start_row + 1 + row_offset, column=col_idx)
                target_cell.value = source_cell.value
    # ----------------------

    return new_end_row

def move_table_and_label_down(ws, table, offset, label_rows_above=1):
    """
    Move a table and its label (above it) down by `offset` rows.
    Also clears styles in the original label cells to avoid format inheritance.
    """
    # Get bounds
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)

    # Compute full range including label
    label_start_row = max(min_row - label_rows_above, 1)
    start_cell = f"{get_column_letter(min_col)}{label_start_row}"
    end_cell = f"{get_column_letter(max_col)}{max_row}"
    full_range = f"{start_cell}:{end_cell}"

    # Store original label cell coordinates
    label_cells_to_clear = [
        (row, col)
        for row in range(label_start_row, min_row)
        for col in range(min_col, max_col + 1)
    ]

    # Move everything
    ws.move_range(full_range, rows=offset, cols=0)

    # Clear styles of old label cells
    for row, col in label_cells_to_clear:
        cell = ws.cell(row=row, column=col)
        cell.font = Font()
        cell.fill = PatternFill()
        cell.border = Border()
        cell.alignment = Alignment()
        cell.number_format = 'General'
        cell.protection = Protection()

    # Update table ref
    new_min_row = min_row + offset
    new_max_row = max_row + offset
    new_start = f"{get_column_letter(min_col)}{new_min_row}"
    new_end = f"{get_column_letter(max_col)}{new_max_row}"
    table.ref = f"{new_start}:{new_end}"

def prepare_tables_for_writing(ws, wb, write_instructions, label_rows_above=1):
    """
    Ensures there's enough space for all tables being written by moving lower tables and their labels down.
    Also moves any named ranges defined in the worksheet that would be overlapped by the table moves.
    
    Parameters:
    - ws: openpyxl worksheet
    - wb: openpyxl workbook (needed to update named ranges)
    - write_instructions: list of tuples (table_name, dataframe)
    - label_rows_above: int, rows reserved above tables for labels
    """
    all_tables = list(ws.tables.values())
    table_map = {t.name: t for t in all_tables}
    table_lengths = {name: len(df) for name, df in write_instructions if name in table_map}

    def table_start_row(table):
        min_col, min_row, _, _ = range_boundaries(table.ref)
        return min_row

    sorted_tables = sorted(all_tables, key=table_start_row)
    new_table_ends = {}

    for idx, table in enumerate(sorted_tables):
        name = table.name
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if name in table_lengths:
            nrows = table_lengths[name]
            new_end_row = min_row + nrows
            new_table_ends[name] = new_end_row
        else:
            new_table_ends[name] = max_row

    for i in range(len(sorted_tables) - 1):
        upper = sorted_tables[i]
        lower = sorted_tables[i + 1]

        upper_end = new_table_ends[upper.name]
        lower_start = table_start_row(lower)

        spacing = 2
        needed_offset = (upper_end + spacing) - lower_start

        if needed_offset > 0:
            move_table_and_label_down(ws, lower, needed_offset, label_rows_above=label_rows_above)
            new_table_ends[lower.name] += needed_offset

            # Move all named ranges in this sheet that start at or below lower_start
            for defined_name in wb.defined_names.values():
                # Only consider ranges in this worksheet
                if defined_name.attr_text is None:
                    continue
                try:
                    sheet_name, cell_range = defined_name.attr_text.split('!')
                except ValueError:
                    continue  # skip invalid formats

                sheet_name = sheet_name.strip("'")
                if sheet_name != ws.title:
                    continue

                # Remove $ for parsing
                cell_range_clean = cell_range.replace('$', '')
                min_col_nr, min_row_nr, max_col_nr, max_row_nr = range_boundaries(cell_range_clean)

                # If the named range starts at or below lower_start, move it
                if min_row_nr >= lower_start:
                    move_named_range(ws, wb, defined_name.name, offset_rows=needed_offset)

def parse_number(value, label):
    if isinstance(value, (int, float)):
        return abs(value)

    if not isinstance(value, str):
        return None

    value = value.strip()
    if not value:
        return None

    # Eliminar cualquier signo negativo y espacios como "- 5.000.000"
    value = value.lstrip('-').replace(' ', '')

    # General cleaning: eliminar símbolos monetarios y letras
    value = re.sub(r'[^\d.,]', '', value)

    # ----- FORMATO POR DOCUMENTO -----
    if label in ['Cuenta BCI 85', 'Cuenta BCI 18', 'Tarjeta de crédito nacional 24', 'Tarjeta de crédito nacional 69']:
        # Ejemplo: 5.050.786 → sin decimales, con separador de miles '.'
        value = value.replace('.', '')
    
    elif label in ['Cuenta BCI Comercio Exterior', 'Tarjeta de crédito internacional', 'Transbank']:
        # Ejemplo: 116.957,02 → miles con '.' y decimales con ','
        value = value.replace('.', '').replace(',', '.')

    elif label == 'Transferencias':
        # Ejemplo: - 5.000.000 → miles con '.', sin decimales, con espacio y signo
        value = value.replace('.', '')

    elif label == 'Siteminder':
        # Ejemplo: 1433.61 → sin separador de miles, decimales con punto
        value = value.replace(',', '')  # por si acaso

    elif label == 'Security':
        # Ejemplo: 2330,51 → sin miles, decimales con coma
        value = value.replace(',', '.')

    else:
        # Fallback: intenta autodetectar
        if re.match(r'^\d{1,3}(\.\d{3})*,\d{1,2}$', value):
            value = value.replace('.', '').replace(',', '.')
        elif ',' in value and '.' not in value:
            value = value.replace(',', '.')
        else:
            value = value.replace(',', '')

    try:
        float_val = float(value)
        return int(float_val) if float_val.is_integer() else float_val
    except ValueError:
        return None

def apply_classification(df, model_path, vectorizer_path, text_columns):
    """
    Apply a classification model to a dataframe using specified text columns.

    Parameters:
        df (pd.DataFrame): The input dataframe.
        model_path (str): Path to the .pkl file for the trained model.
        vectorizer_path (str): Path to the .pkl file for the vectorizer.
        text_columns (list): List of column names whose values are used as input to the model.

    Returns:
        pd.DataFrame: The original dataframe with two new columns:
                      ['Predicción Clave', 'Puntaje Confianza']
    """
    # Load model and vectorizer
    with open(model_path, "rb") as model_file, open(vectorizer_path, "rb") as vectorizer_file:
        model = pickle.load(model_file)
        vectorizer = pickle.load(vectorizer_file)

    # Ensure only available columns are used
    available_cols = [col for col in text_columns if col in df.columns]

    def classify_row(row):
        inputs = [str(row[col]) for col in available_cols if pd.notnull(row[col])]
        return pd.Series(classify_text(model, vectorizer, *inputs))

    df[['CLAVE', 'PREDICCIÓN']] = df.apply(classify_row, axis=1)
    return df


def file_to_df(path, label):
    if label == 'Banco Security':
        header = 14
        df = pd.read_excel(path, header=header - 1, thousands='.')
        end_index = None
        # Itera hasta encontrar un espacio y para
        for i, row in df.iterrows():
            if pd.isna(row[0]):
                end_index = i
                break 
        df = df.iloc[: end_index]

    elif label == 'Transbank':
        df = process_transbank_directory(path)
    
    elif label == 'Siteminder':
        df = pd.read_csv(path)

    elif label in ('Cuenta BCI Comercio Exterior', 'Tarjeta de crédito nacional 24', 'Tarjeta de crédito nacional 69', 'Tarjeta de crédito internacional'):
        # Aquí se abre tu GUI personalizada que devuelve un DataFrame
        df = launch_pdf_gui(path, label)  # Esta GUI procesa el PDF y devuelve df

    elif label == 'Cuenta BCI 18':
        header = 18 # Fila donde están los nombres de la columna
        df = pd.read_excel(path, header=header - 1, thousands='.')

    elif label == 'Cuenta BCI 85':
        header = 23 # Fila donde están los nombres de la columna (Numero de fila en excel - 1)
        df = pd.read_excel(path, header=header - 1, thousands='.')

    elif label == 'Transferencias':
        engine = 'xlrd' # El motor openpyxl no usa xls
        header = 9 # Fila donde están los nombres de la columna (Numero de fila en excel - 1)
        df = pd.read_excel(path, header=header - 1, thousands='.')

    df = df[columns_to_extract[label]]
    df.rename(columns=column_dictionary[label], inplace=True)
    df = format_column(df, label)

    return df

def convert_numeric(value):
    value = value.replace('.', '')
    numeric = float(value.split()[1])
    return numeric

def format_column(df, label):
    if label in ['Cuenta BCI 85', 'Cuenta BCI 18']:
        numeric_cols = ['CARGO', 'ABONO', 'SALDO']
        for col in numeric_cols:
            df[col] = df[col].apply(lambda x: parse_number(x, label))

    elif label == 'Cuenta BCI Comercio Exterior':
        numeric_cols = ['CARGO', 'ABONO', 'SALDO']
        for col in numeric_cols:
            df[col] = df[col].apply(lambda x: parse_number(x, label))

    elif label == 'Transferencias':
        numeric_cols = ['CARGO TRANSF']
        for col in numeric_cols:
            df[col] = df[col].apply(lambda x: parse_number(x, label))

    elif label in ['Tarjeta de crédito nacional 24', 'Tarjeta de crédito nacional 69']:
        numeric_cols = ['CARGO']
        for col in numeric_cols:
            df[col] = df[col].apply(lambda x: x.split()[-1] if isinstance(x, str) else x)
            df[col] = df[col].apply(lambda x: parse_number(x, label))

    elif label == 'Tarjeta de crédito internacional':
        numeric_cols = ['MONTO US$']
        for col in numeric_cols:
            df[col] = df[col].apply(lambda x: parse_number(x, label))

    elif label == 'Siteminder':
        # Formato numérico
        numeric_cols = ['TOTAL']
        for col in numeric_cols:
            df[col] = df[col].apply(lambda x: x.split()[-1] if isinstance(x, str) else x)
            df[col] = df[col].apply(lambda x: parse_number(x, label))
        # Limpiar y convertir columnas
        df['HABITACIONES'] = df['HABITACIONES'].apply(lambda x: x[0])
        df['LLEGADA'] = df['LLEGADA'].apply(convert_date).dt.date
        df['SALIDA'] = df['SALIDA'].apply(convert_date).dt.date

    elif label == 'Security':
        numeric_cols = ['CARGO', 'ABONO', 'SALDO']
        for col in numeric_cols:
            df[col] = df[col].apply(lambda x: parse_number(x, label))

    elif label == 'Transbank':

        # Nuevas columnas
        df['PESOS'] = np.nan
        df['TIPO VENTA'] = np.nan
        df['PROPINA'] = np.nan
        df['CODIGO EMPLEADO'] = np.nan

        # Asignar moneda
        df['MONEDA'] = df['TIPO MOVIMIENTO'].replace({
            'Venta USD$': 'DOLAR',
            'Venta $': 'PESO'
        })

        # Convertir número de boleta
        df['NUMERO'] = pd.to_numeric(df['NUMERO BOLETA'], errors='coerce')

        # Convertir fecha y ordenar
        df[['FECHA', 'HORA']] = df['FECHA'].str.split(' ', n=1, expand=True)
        df['FECHA'] = pd.to_datetime(df['FECHA'], format='%d/%m/%Y', errors='coerce')
        df = df.sort_values(by=['FECHA', 'NUMERO'], ascending=True)

    return df

def process_transbank_directory(directory):
    dataframes = []
    for filename in os.listdir(directory):
        if not filename.lower().endswith(('.xlsx', '.xls')):
            continue

        file_path = os.path.join(directory, filename)

        try:
            df_temp = pd.read_excel(file_path, header=None)
            columnas = df_temp.iloc[17].astype(str).str.strip() + " - " + df_temp.iloc[18].astype(str).str.strip()

            df = pd.read_excel(file_path, skiprows=19, header=None)
            df.columns = columnas

            dataframes.append(df)

        except Exception as e:
            print(f"Error processing {filename}: {e}")
            continue

    if dataframes:
        df_stack = pd.concat(dataframes, ignore_index=True)
        return df_stack
    else:
        return pd.DataFrame()

def df_to_wb(dfs, wb, sheet_name):
    ws = wb[sheet_name]
    table_data = []  # Will contain tuples (table_name, df)
    if sheet_name == 'BCI':
        # ------------------ Cuenta BCI 85 + Transferencias ------------------
        if 'Transferencias' in dfs and 'Cuenta BCI 85' in dfs:
            df_transfer = dfs['Transferencias'].copy()
            df_85 = dfs['Cuenta BCI 85'].copy()
            desc_list = ['TRASPASO FONDOS OTRO BANCO EN LINEA', 'CARGO POR TRANSF DE FONDOS AUTOSERVICIO']
            mask_desc = df_85['MOVIMIENTO'].isin(desc_list) | df_85['MOVIMIENTO'].str.contains('Transfer', case=False, na=False)
            filtered_cartola = df_85[mask_desc].copy()
            filtered_cartola['merge_id'] = filtered_cartola.groupby(['CARGO']).cumcount()
            df_transfer['merge_id'] = df_transfer.groupby(['CARGO TRANSF']).cumcount()

            merged_cartola = filtered_cartola.merge(
                df_transfer[['CARGO TRANSF', 'NOMBRE DESTINO', 'MENSAJE', 'merge_id']],
                how='left',
                left_on=['CARGO', 'merge_id'],
                right_on=['CARGO TRANSF', 'merge_id']
            )

            non_filtered_cartola = df_85[~mask_desc].copy()
            df_final = pd.concat([merged_cartola, non_filtered_cartola]).sort_index()
            model_path, vectorizer_path = join(PATH_MODEL, "modelo_C85.pkl"), join(PATH_MODEL, "vectorizer_C85.pkl")
            df_final = apply_classification(df_final, model_path, vectorizer_path, ['MOVIMIENTO', 'NOMBRE DESTINO', 'MENSAJE'])
            df_final = df_final.drop(columns=['CARGO TRANSF', 'merge_id'], errors='ignore')
            table_data.append(('Cuenta_85', df_final))

        # ------------------ Cuenta BCI Comercio Exterior ------------------
        if 'Cuenta BCI Comercio Exterior' in dfs:
            df_ext = dfs['Cuenta BCI Comercio Exterior'].copy()
            model_path, vectorizer_path = join(PATH_MODEL, "modelo_CUSD.pkl"), join(PATH_MODEL, "vectorizer_CUSD.pkl")
            df_ext = apply_classification(df_ext, model_path, vectorizer_path, ['DESCRIPCION'])
            table_data.append(('BCI_Comercio_Exterior', df_ext))

        # ------------------ Tarjeta de crédito nacional ------------------
        if 'Tarjeta de crédito nacional 24' in dfs or 'Tarjeta de crédito nacional 69' in dfs:
            df_tcn_list = [dfs[k] for k in ['Tarjeta de crédito nacional 24', 'Tarjeta de crédito nacional 69'] if k in dfs]
            df_credito_nacional = pd.concat(df_tcn_list, ignore_index=True)
            model_path, vectorizer_path = join(PATH_MODEL, "modelo_TCN.pkl"), join(PATH_MODEL, "vectorizer_TCN.pkl")
            df_credito_nacional = apply_classification(df_credito_nacional, model_path, vectorizer_path, ['DESCRIPCION'])
            table_data.append(('TC_Nacional', df_credito_nacional))  # Use the actual table name here

        # ------------------ Tarjeta de crédito internacional ------------------
        if 'Tarjeta de crédito internacional' in dfs:
            df_credito_internacional = dfs['Tarjeta de crédito internacional'].copy()
            model_path, vectorizer_path = join(PATH_MODEL, "modelo_TCI.pkl"), join(PATH_MODEL, "vectorizer_TCI.pkl")
            df_credito_internacional = apply_classification(df_credito_internacional, model_path, vectorizer_path, ['DESCRIPCION'])
            df_credito_internacional['DESCRIPCION'] = df_credito_internacional['CIUDAD'].fillna('') + ' ' + df_credito_internacional['DESCRIPCION'].fillna('')
            start_row = 1  # Dummy value for now
            df_credito_internacional['CARGO'] = [
                f"=D{row_num}*EERR!$D$2" for row_num in range(start_row, start_row + len(df_credito_internacional))
            ]
            table_data.append(('TC_Internacional', df_credito_internacional))  # Use actual table name

    elif sheet_name == 'Security':
        if 'Banco Security' in dfs:
            df_security = dfs['Banco Security']
            model_path, vectorizer_path = join(PATH_MODEL, "modelo_security.pkl"), join(PATH_MODEL, "vectorizer_security.pkl")
            text_columns = ['DESCRIPCION']
            df_security = apply_classification(df_security, model_path, vectorizer_path, text_columns)
            table_data.append(('Security', df_security))
    elif sheet_name == 'Siteminder':
        if 'Siteminder' in dfs:
            df_siteminder = dfs['Siteminder']
            table_data.append(('Siteminder', df_siteminder))
    elif sheet_name == 'BCI FondRendir':
        if 'Cuenta BCI 18' in dfs:
            df_cta18 = dfs['Cuenta BCI 18']
            table_data.append(('Cuenta_18', df_cta18))
    elif sheet_name == 'Transbank':
        if 'Transbank' in dfs:
            df_trans = dfs['Transbank']
            print(len(df_trans))
            table_data.append(('Transbank', df_trans))

    # ------------------ Apply all changes to sheet ------------------
    if table_data:
        prepare_tables_for_writing(ws, wb, table_data)
        for table_name, df in table_data:
            write_df_to_named_table_by_header(ws, sheet_name, table_name, df)

def launch_pdf_gui(file_path, file_type=None):
    import tkinter as tk
    import pandas as pd
    root = tk.Toplevel()
    extractor = TableExtractorApp(root, file_path=file_path)
    root.grab_set()
    root.wait_window()
    return getattr(extractor, 'result_df', pd.DataFrame())

