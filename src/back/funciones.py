import openpyxl

archivo_excel = "./src/inventario.xlsx"

def listar_categorias(archivo_excel):
    try:
        libro = openpyxl.load_workbook(archivo_excel)
        categorias = libro.sheetnames
        print("Las categorias son: ", categorias)
        return categorias
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return []

def listar_productos(categoria, archivo_excel):
    try:
        libro = openpyxl.load_workbook(archivo_excel)
        if categoria not in libro.sheetnames:
            print("Categoria no encontrada")
            return []

        hoja = libro[categoria]
        productos = []
        for fila in hoja.iter_rows(min_row=1, max_row=hoja.max_row, max_col=hoja.max_column):
            for celda in fila:
                if celda.value == "DETALLE":
                    columna_detalle = celda.column
                    for producto_celda in hoja.iter_rows(min_row=celda.row + 1, max_col=columna_detalle, max_row=hoja.max_row):
                        producto = producto_celda[columna_detalle - 1].value
                        if producto:
                            productos.append(producto)
                    return productos
        print("No se encontraron productos en la categoria")
        return productos
    except Exception as e:
        print(f"Error al listar productos: {e}")
        return []

def buscar_codigo_producto(categoria, nombre_producto, archivo_excel, columnas_codigo=1):
    try:
        libro = openpyxl.load_workbook(archivo_excel)
        if categoria not in libro.sheetnames:
            print("Categoria no encontrada")
            return None

        hoja = libro[categoria]
        for fila in hoja.iter_rows(min_row=1, max_row=hoja.max_row, max_col=hoja.max_column):
            for celda in fila:
                if celda.value == "DETALLE":
                    columna_detalle = celda.column
                    for producto_celda in hoja.iter_rows(min_row=celda.row + 1, max_col=columna_detalle, max_row=hoja.max_row):
                        if producto_celda[columna_detalle - 1].value == nombre_producto:
                            codigo_celda = producto_celda[columna_detalle - 1].offset(column=columnas_codigo)
                            columna_codigo = codigo_celda.column
                            fila_numero = codigo_celda.row
                            columna_letra = openpyxl.utils.get_column_letter(columna_codigo)
                            print(f"El código del producto '{nombre_producto}' es: {codigo_celda.value} (celda {columna_letra}{fila_numero})")
                            return codigo_celda.value, columna_letra, fila_numero
        print(f"No se encontró el producto '{nombre_producto}' en la categoría '{categoria}'")
        return None
    except Exception as e:
        print(f"Error al buscar código de producto: {e}")
        return None

def while_loop(archivo_excel):
    while True:
        categoria = input("Digite la categoria del producto: ")
        nombre_producto = input("Digite el nombre del producto: ")
        columnas_codigo = int(input("Ingrese el número de columnas a la derecha de 'DETALLE' donde se encuentra el código del producto: "))

        codigo_producto = buscar_codigo_producto(categoria, nombre_producto, archivo_excel, columnas_codigo)

        if codigo_producto is not None:
            print(f"El código del producto '{nombre_producto}' es: {codigo_producto}")
        else:
            print(f"No se encontró el producto '{nombre_producto}' en el archivo.")

        resp = input("¿Desea continuar? (si/no): ").strip().lower()
        if resp != "si":
            break

if __name__ == '__main__':
    listar_categorias(archivo_excel)
    while_loop(archivo_excel)
    print("Fin del programa")
