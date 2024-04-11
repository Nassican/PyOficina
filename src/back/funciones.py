import openpyxl
archivo_excel = ("./src/inventario.xlsx")


def listar_categorias(archivo_excel):
    # Nombres de las hojas de excel
    libro = openpyxl.load_workbook(archivo_excel)
    print("Las categorias son: ", libro.sheetnames)
    return libro.sheetnames


def listar_productos(categoria, archivo_excel):
    productos = []

    libro = openpyxl.load_workbook(archivo_excel)

    # Ir a la hoja de excel de la categoria
    if categoria in libro.sheetnames:
        hoja = libro[categoria]
        # Buscar la fila de encabezados que contiene "Productos"
        for columna in range(1, hoja.max_column + 1):
            if hoja.cell(row=1, column=columna).value == "Productos":
                # Iterar sobre las filas para obtener los nombres de los productos
                for fila in range(2, hoja.max_row + 1):
                    producto = hoja.cell(row=fila, column=columna).value
                    if producto:
                        productos.append(producto)
                break  # Salir del bucle una vez que se encuentre la columna de productos
    else:
        print("Categorias terminadas, no encontradas o productos no encontrados o no hay productos en la categoria")

    return productos


def buscar_codigo_producto(categoria, nombre_producto, archivo_excel):

    libro = openpyxl.load_workbook(archivo_excel)

    # Ir a la hoja de excel de la categoria
    if categoria in libro.sheetnames:
        hoja = libro[categoria]
        for columna in range(1, hoja.max_column + 1):
            if hoja.cell(row=1, column=columna).value == "Productos":
                for fila in range(1, hoja.max_row + 1):
                    for columna in range(1, hoja.max_column + 1):
                        if nombre_producto == hoja.cell(row=fila, column=columna).value:
                            # Retorno del código del producto
                            return hoja.cell(row=fila, column=columna + 1).value
    else:
        print("Categoria no encontrada")
    return None


def while_loop():
    resp = "si"
    while resp == "si":

        categoria = input("Digite la categoria del producto: ")
        nombre_producto = input("Digite el nombre del producto: ")

        codigo_producto = buscar_codigo_producto(
            categoria, archivo_excel, nombre_producto)

        if codigo_producto is not None:
            print(f"El código del producto '{nombre_producto}' es: {codigo_producto}")
        else:
            print(f"No se encontró el producto '{nombre_producto}' en el archivo.")

        resp = input("¿Desea continuar? (si/no): ")


if __name__ == '__main__':
    listar_categorias(archivo_excel)
    print("Fin del programa")
