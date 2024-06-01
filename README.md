# PyOficina

Gestor de productos y codigos de un excel con interfaz grafica

## Instrucciones de su uso:

### Instalar dependencias:

Para realizar el correcto funcionamiento se necesitan las librerias PySide6 y openpyxl, la puedes instalar desde pip con el siguiente comando:

```bash
  pip install PySide6
  pip install openpyxl
```
o si tienes el archivo **requirements.txt** puedes instalar todas las dependencias con el siguiente comando:

```bash
  pip install -r requirements.txt
```

### Despliegue de la aplicacion

Ejecuta el archivo **main.py** con tu interprete de python

```bash
  python main.py
```

### Formato del archivo excel para el correcto funcionamiento

El archivo excel debe tener la siguiente estructura:

Por defecto el archivo buscara la columna con el nombre **DETALLE** para poder realizar la busqueda de los productos y el codigo por defecto tomara la siguiente columna a la derecha de la columna **DETALLE**.

**Ejemplo:**

| DETALLE | {Codigo} |
|--------|--------|
| Producto 1 | 0001 |