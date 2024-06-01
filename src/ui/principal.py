import sys
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QComboBox,
    QLineEdit,
    QVBoxLayout,
    QPushButton,
    QDialog,
    QHBoxLayout,
    QFileDialog,
    QMessageBox,
    QCompleter
)
from PySide6.QtCore import Qt, QSortFilterProxyModel
from PySide6.QtGui import QFont, QFontDatabase, QIcon, QPixmap
import os
from ..back import funciones as excel

# Importar el archivo de recursos generado
import recursos2_rc

# Ruta de este archivo
path = os.path.dirname(os.path.abspath(__file__))
window_icon_path = os.path.join(path, "img/escudo.png")

# Colores y estilos
GREEN_COLOR = "#00FF00"
GREEN_COLOR_DARK = "#007D00"
BANNER_STYLE = """
    QLabel {
    background-color: green;
    color: white;
    font-size: 40px;
    font-weight: bold;
    padding: 10px;
    }
"""
BUTTON_STYLE = """
    QPushButton {
    font-size: 20px;
    background-color: green;
    color: white;
    border: 1px solid gray;
    border-radius: 10px;
    padding: 5px;
    }
    QPushButton:hover {
    background-color: lightgreen;
    color: gray;
    }
    QPushButton:pressed {
    background-color: darkgreen;
    }
"""
BUTTON_STYLE_DIABLED = """
    QPushButton {
    font-size: 20px;
    background-color: lightgray;
    color: gray;
    border: 1px solid gray;
    border-radius: 10px;
    padding: 5px;
    }
"""
LABEL_STYLE = """
    QLabel {
    font-size: 20px;
    font-weight: bold;
    }
"""
COMBO_BOX_STYLE = """
    QComboBox {
    background-color: white;
    color: black;
    border: 1px solid gray;
    border-radius: 10px;
    font-size: 20px;
    }
    QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 50px;
    border-left-width: 1px;
    border-left-color: darkgray;
    border-left-style: solid;
    border-top-right-radius: 5px;
    border-bottom-right-radius: 5px;
    }
    QComboBox::down-arrow {
        image: url(:/down-arrow.png)
    }
    QComboBox QAbstractItemView {
        background-color: white;
        border: 1px solid gray;
        border-radius: 5px;
        padding: 5px;
        color: black;
    }
    QComboBox QAbstractItemView::item {
        padding: 5px;
    }
    QComboBox QAbstractItemView::item:selected {
        background-color: lightgray;
        color: black;
    }
"""
# ComboBox y Labels estilo
COMBO_BOX_STYLE_DISABLED = """
    QComboBox {
    background-color: lightgray;
    color: gray;
    border: 1px solid gray;
    border-radius: 10px;
    padding: 5px;
    font-size: 20px;
    }
    QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    padding: 5px;
    border-left-width: 0px;
    border-left-color: darkgray;
    border-left-style: solid; /* just a single line */
    border-top-right-radius: 3px; /* same radius as the QComboBox */
    border-bottom-right-radius: 3px;
    }
    QComboBox QAbstractItemView {
    background-color: white;
    border: 1px solid gray;
    border-radius: 10px;
    }
"""
INPUT_STYLE = """
    QLineEdit {
    font-size: 20px;
    background-color: white;
    color: black;
    border: 1px solid gray;
    border-radius: 10px;
    padding: 5px;
    }
"""
INPUT_STYLE_DISABLED = """
    QLineEdit {
    font-size: 20px;
    background-color: lightgray;
    color: gray;
    border: 1px solid gray;
    border-radius: 10px;
    padding: 5px;
    }
"""


def cargar_fuentes():
    QFontDatabase.addApplicationFont(os.path.join(path, "fonts/RedHatText-Bold.ttf"))
    QFontDatabase.addApplicationFont(os.path.join(path, "fonts/RedHatText-Regular.ttf"))
    QFontDatabase.addApplicationFont(os.path.join(path, "fonts/Overpass-Bold.ttf"))

class MessageWidget(QMessageBox):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Oficina de Planeación")
        self.setWindowIcon(QIcon(window_icon_path))
        self.setStyleSheet("""
            QMessageBox {
            border: 10px;
            font-size: 20px;
            }
        """)
        self.button_message = self.addButton("Aceptar", QMessageBox.AcceptRole)
        self.button_message.setCursor(Qt.PointingHandCursor)
        self.setIcon(QMessageBox.Critical)
        self.button_message.setStyleSheet(BUTTON_STYLE)

class VentanaPrincipal(QWidget):


    pathToFileComplete = ""

    def __init__(self):
        super().__init__()
        self.initUI()
        self.setWindowIcon(QIcon(window_icon_path))
        self.messageBox = MessageWidget()
        self.editando = False
        self.añadiendo = False
        self.setStyleSheet("background-color: white; color: black;")

    def mostrar_info_emergente(self):
        if not self.pathToFileComplete:
            self.messageBox.setText("No hay archivo seleccionado")
            self.messageBox.exec()
            return

        self.buscar_codigo_producto()

    def cargar_categorias(self):
        categorias = excel.listar_categorias(self.pathToFileComplete)
        categorias.sort()
        self.categoria_combo.clear()
        self.categoria_combo.addItems(categorias)

    def onChangedCategoria(self):
        categoria = self.categoria_combo.currentText()
        productos = excel.listar_productos(categoria, self.pathToFileComplete)
        print(f"Los productos de la categoria {categoria} son: {len(productos)}")
        self.cantidad_productos.setText(f"({len(productos)} productos)")
        self.producto_combo.clear()
        self.producto_combo.addItems(productos)
        self.producto_proxy_model.setSourceModel(self.producto_combo.model())
        self.verificar_habilitar_campos_añadir()

        # Mensaje de alerta en caso de no encontrar ningun producto
        if len(productos) == 0:
            self.label_mensajes_alerta.setStyleSheet("font-size: 15px; font-weight: bold; color: red; margin-left: 30px; margin-right: 30px;")
            self.label_mensajes_alerta.setText("No se encontraron productos en la categoría seleccionada")
        else:
            self.label_mensajes_alerta.setText("")

    def on_producto_text_changed(self, text):
        self.producto_proxy_model.setFilterFixedString(text)

    def buscar_codigo_producto(self):
        categoria_texto = self.categoria_combo.currentText()
        producto_texto = self.producto_combo.currentText()
        try:
            codigo, self.columna_letra_codigo, self.fila_numero_codigo, self.columna_letra_producto, self.fila_numero_producto = excel.buscar_codigo_producto(categoria_texto, producto_texto, self.pathToFileComplete)
            celda_codigo = f"{self.columna_letra_codigo}{self.fila_numero_codigo}"
            self.varTextoNomProd = producto_texto
            self.varTextTipoProd = categoria_texto
            self.varNumCodProd = f"{codigo} (celda {celda_codigo})"
            
            self.nombre_seleccion.setText(self.varTextoNomProd)
            self.nombre_seleccion.setCursor(Qt.IBeamCursor)
            self.categoria_seleccion.setText(self.varTextTipoProd)
            self.codigo_seleccion.setText(str(codigo))
            self.celda_seleccion.setText(celda_codigo)
            self.verificar_habilitar_campos_edicion()
            self.verificar_habilitar_campos_añadir()
            self.desactivar_añadir()
            self.desactivar_edicion()
            self.inhabilitar_campos()
        except Exception as e:
            self.messageBox.setText("No se encontró el producto en el archivo")
            self.messageBox.exec()

    # Funcion para habilitar los LineEdit para poderlos editar
    def habilitar_campos_edicion(self):
        if not self.editando:
            self.alternar_botones("editar")
            self.nombre_seleccion.setReadOnly(False)
            self.codigo_seleccion.setReadOnly(False)
            self.nombre_seleccion.setStyleSheet(INPUT_STYLE)
            self.codigo_seleccion.setStyleSheet(INPUT_STYLE)
            self.codigo_seleccion.setCursor(Qt.IBeamCursor)
            self.nombre_seleccion.setCursor(Qt.IBeamCursor)
            self.button_edit.setText("Guardar")
            self.editando = True
            self.desactivar_añadir()
        else:
            self.nombre_seleccion.setReadOnly(True)
            self.codigo_seleccion.setReadOnly(True)
            self.nombre_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
            self.codigo_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
            self.nombre_seleccion.setCursor(Qt.IBeamCursor)
            self.codigo_seleccion.setCursor(Qt.IBeamCursor)
            self.guardar_edicion()
            self.desactivar_edicion()

    def habilitar_campos_añadir(self):
        if not self.añadiendo:
            self.alternar_botones("añadir")
            self.nombre_seleccion.setReadOnly(False)
            self.codigo_seleccion.setReadOnly(False)
            self.nombre_seleccion.setStyleSheet(INPUT_STYLE)
            self.codigo_seleccion.setStyleSheet(INPUT_STYLE)
            self.codigo_seleccion.setCursor(Qt.IBeamCursor)
            self.nombre_seleccion.setCursor(Qt.IBeamCursor)
            self.categoria_seleccion.setText(self.categoria_combo.currentText())
            self.nombre_seleccion.setText("")
            self.codigo_seleccion.setText("")
            self.celda_seleccion.setText("")
            self.button_añadir.setText("Guardar")
            self.añadiendo = True
            self.desactivar_edicion()
        else:
            self.agregar_producto()
            self.nombre_seleccion.setReadOnly(True)
            self.codigo_seleccion.setReadOnly(True)
            self.nombre_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
            self.codigo_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
            self.nombre_seleccion.setCursor(Qt.IBeamCursor)
            self.codigo_seleccion.setCursor(Qt.IBeamCursor)
            self.desactivar_añadir()

    def desactivar_edicion(self):
        self.button_edit.setText("Editar")
        self.editando = False

    def desactivar_añadir(self):
        self.button_añadir.setText("Añadir Producto")
        self.añadiendo = False

    def inhabilitar_campos(self):
        self.nombre_seleccion.setReadOnly(True)
        self.codigo_seleccion.setReadOnly(True)
        self.nombre_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
        self.codigo_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
        self.nombre_seleccion.setCursor(Qt.IBeamCursor)
        self.codigo_seleccion.setCursor(Qt.IBeamCursor)

    def verificar_habilitar_campos_añadir(self):
        # Verificar si hay un archivo seleccionado
        pathToFileComplete = self.pathToFileComplete
        if pathToFileComplete == "":
            self.button_añadir.setEnabled(False)
            self.button_añadir.setStyleSheet(BUTTON_STYLE_DIABLED)
            return
        elif self.categoria_combo.currentText() == "Seleccione una categoría":
            self.button_añadir.setEnabled(False)
            self.button_añadir.setStyleSheet(BUTTON_STYLE_DIABLED)
            return
        elif pathToFileComplete != "" and self.categoria_combo.currentText() != "Seleccione una categoría":
            self.button_añadir.setEnabled(True)
            self.button_añadir.setStyleSheet(BUTTON_STYLE)
            return

    def alternar_botones(self, boton_activado):
        if boton_activado == "añadir":
            self.button_edit.setEnabled(False)
            self.button_edit.setStyleSheet(BUTTON_STYLE_DIABLED)
            self.button_añadir.setEnabled(True)
            self.button_añadir.setStyleSheet(BUTTON_STYLE)
        elif boton_activado == "editar":
            self.button_edit.setEnabled(True)
            self.button_edit.setStyleSheet(BUTTON_STYLE)
            self.button_añadir.setEnabled(False)
            self.button_añadir.setStyleSheet(BUTTON_STYLE_DIABLED)

    def verificar_habilitar_campos_edicion(self):
        if self.nombre_seleccion.text() == "" or self.codigo_seleccion.text() == "":
            self.button_edit.setEnabled(False)
            self.button_edit.setStyleSheet(BUTTON_STYLE_DIABLED)
        else:
            self.button_edit.setEnabled(True)
            self.button_edit.setStyleSheet(BUTTON_STYLE)

    def guardar_edicion(self):
        try:
            # Obtener los valores de los campos de edición
            nuevo_nombre = self.nombre_seleccion.text()
            codigo_producto = self.codigo_seleccion.text()
            if nuevo_nombre == "":
                self.messageBox.setText("El nombre del producto no puede estar vacío")
                self.messageBox.exec()
                return
            if codigo_producto == "":
                self.messageBox.setText("El código del producto no puede estar vacío")
                self.messageBox.exec()
                return
            if not codigo_producto.isdigit():
                self.messageBox.setText("El código del producto debe ser un número")
                self.messageBox.exec()
                return

            # Obtener la información de la celda donde se encuentra el producto
            fila_numero_producto = self.fila_numero_producto
            columna_letra_producto = self.columna_letra_producto

            # Obtener la información de la celda donde se encuentra el código del producto
            fila_numero_codigo = self.fila_numero_codigo
            columna_letra_codigo = self.columna_letra_codigo
            
            categoria = self.categoria_combo.currentText()

            # Guardar la edición en el archivo Excel
            excel.guardar_edicion_producto(
                self.pathToFileComplete,
                nuevo_nombre,
                int(codigo_producto),
                categoria,
                fila_numero_producto,
                columna_letra_producto,
                fila_numero_codigo,
                columna_letra_codigo
            )

            # Mostrar mensaje de éxito
            self.messageBox.setIcon(QMessageBox.Information)
            self.messageBox.setText("Producto actualizado correctamente")
            self.messageBox.exec()

            # Actualizar los campos de productos
            self.recargar_documento()

            message=f"Se guardó la edición del producto '{nuevo_nombre}' con código '{codigo_producto}' en la categoría '{categoria}'"
            self.label_mensajes_alerta.setText(message)
            self.label_mensajes_alerta.setStyleSheet("font-size: 15px; font-weight: bold; color: #007D00; margin-left: 30px; margin-right: 30px;")
            # Reiniciar campos
            self.resetear_campos()
            self.verificar_habilitar_campos_edicion()
        except Exception as e:
            # Mostrar mensaje de error si ocurre alguna excepción
            self.messageBox.setIcon(QMessageBox.Critical)
            self.messageBox.setText(f"Error al guardar edición: {e}")
            self.messageBox.exec()

    def recargar_documento(self):
        self.cargar_categorias()
        self.onChangedCategoria()

    def resetear_campos(self):
        self.nombre_seleccion.setText("")
        self.categoria_seleccion.setText("")
        self.codigo_seleccion.setText("")
        self.celda_seleccion.setText("")

    def agregar_producto(self):
        try:
            # Obtener los valores de los campos de edición
            nuevo_nombre = self.nombre_seleccion.text()
            codigo_producto = self.codigo_seleccion.text()
            if nuevo_nombre == "":
                self.messageBox.setText("El nombre del producto no puede estar vacío")
                self.messageBox.exec()
                return
            if codigo_producto == "":
                self.messageBox.setText("El código del producto no puede estar vacío")
                self.messageBox.exec()
                return
            if not codigo_producto.isdigit():
                self.messageBox.setText("El código del producto debe ser un número")
                self.messageBox.exec()
                return

            categoria = self.categoria_combo.currentText()

            # Guardar la edición en el archivo Excel
            excel.agregar_producto(
                self.pathToFileComplete,
                categoria,
                nuevo_nombre,
                int(codigo_producto)
            )

            # Mostrar mensaje de éxito
            self.messageBox.setIcon(QMessageBox.Information)
            self.messageBox.setText("Producto añadido correctamente")
            self.messageBox.exec()

            # Actualizar los campos de productos
            self.recargar_documento()

            message=f"Se añadió el producto '{nuevo_nombre}' con código '{codigo_producto}' en la categoría '{categoria}'"
            self.label_mensajes_alerta.setText(message)
            self.label_mensajes_alerta.setStyleSheet("font-size: 15px; font-weight: bold; color: #007D00; margin-left: 30px; margin-right: 30px;")
            # Reiniciar campos
            self.resetear_campos()
            self.verificar_habilitar_campos_edicion()
        except Exception as e:
            # Mostrar mensaje de error si ocurre alguna excepción
            self.messageBox.setIcon(QMessageBox.Critical)
            self.messageBox.setText(f"Error al añadir producto: {e}")
            self.messageBox.exec()

    def ventana_escoger_archivo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setNameFilter("Microsoft Excel (*.xlsx)")
        file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

        maximosCaracteres = 50

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                pathToFileComplete = selected_files[0]
                pathToFileLabel = pathToFileComplete
                # Texto limitado desde el final
                if len(pathToFileLabel) > maximosCaracteres:
                    pathToFileLabel = " . . . " + \
                        pathToFileLabel[-maximosCaracteres:]

                self.pathToFileComplete = pathToFileComplete
                self.labelPathToFile.setText(pathToFileLabel)
                self.checkear_ruta_del_archivo()
                self.cargar_categorias()
                self.onChangedCategoria()
                self.resetear_campos()
                print(pathToFileComplete)

    def checkear_ruta_del_archivo(self):
        # Verificar si hay un archivo seleccionado
        pathToFileComplete = self.pathToFileComplete

        if pathToFileComplete != "":
            # Activar inputs y combo
            self.categoria_combo.setEnabled(True)
            self.categoria_combo.setStyleSheet(COMBO_BOX_STYLE)
            self.producto_combo.setEnabled(True)
            self.producto_combo.setStyleSheet(COMBO_BOX_STYLE)
            self.categoria_combo.setMinimumWidth(self.width() - 200)
            self.producto_combo.setMinimumWidth(self.width() - 200)

            print(f"Archivo seleccionado, RUTA: {pathToFileComplete}")
        else:
            # Desactivar inputs y combo
            self.categoria_combo.setEnabled(False)
            self.categoria_combo.setStyleSheet(COMBO_BOX_STYLE_DISABLED)
            self.producto_combo.setEnabled(False)
            self.producto_combo.setStyleSheet(COMBO_BOX_STYLE_DISABLED)
            self.categoria_combo.setMinimumWidth(self.width() - 200)
            self.producto_combo.setMinimumWidth(self.width() - 200)
            print(f"No hay archivo seleccionado, RUTA: {pathToFileComplete}")

    def initUI(self):
        self.setWindowTitle("Oficina de Planeación")
        self.resize(900, 550)
        self.setMaximumSize(900, 600)
        self.setMinimumSize(900, 550)

        fuente = QFont("Red Hat Text", 20)

        # Banner superior verde
        banner = QLabel("Oficina de Planeación", self)
        banner.setMaximumHeight(100)
        banner.setMinimumWidth(445)
        banner.setFont(fuente)
        banner.setStyleSheet(BANNER_STYLE)


        # Layout principal
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Layout para formulario
        form_layout = QVBoxLayout()
        form_layout.setAlignment(Qt.AlignCenter)
        form_layout.setContentsMargins(100, 20, 100, 20)
        form_layout.maximumSize()

        # Categoría
        categoria_label = QLabel("Categoría:", self)
        categoria_label.setStyleSheet(LABEL_STYLE)
        categoria_label.setFont(fuente)
        self.categoria_combo = QComboBox(self)
        self.categoria_combo.setFont(fuente)
        self.categoria_combo.setStyleSheet(COMBO_BOX_STYLE)
        self.categoria_combo.setEnabled(False)
        self.categoria_combo.addItem("Seleccione una categoría")
        self.categoria_combo.currentTextChanged.connect(self.onChangedCategoria)
        self.categoria_combo.setMinimumHeight(40)

        form_layout.addWidget(categoria_label)
        form_layout.addWidget(self.categoria_combo)

        # Producto
        producto_num_productos = QHBoxLayout()
        producto_num_productos.setContentsMargins(0, 0, 0, 0)

        producto_label = QLabel("Producto:", self)
        producto_label.setStyleSheet(LABEL_STYLE)
        producto_label.setFont(fuente)
        #
        self.cantidad_productos = QLabel("", self)
        # Alinearlo a la derecha
        self.cantidad_productos.setAlignment(Qt.AlignRight)
        self.cantidad_productos.setFont(fuente)
        self.cantidad_productos.setStyleSheet("font-size: 20px; color: gray;")
        producto_num_productos.addWidget(producto_label)
        producto_num_productos.addWidget(self.cantidad_productos)

        self.producto_combo = QComboBox(self)
        self.producto_combo.setFont(fuente)
        self.producto_combo.setStyleSheet(COMBO_BOX_STYLE)
        self.producto_combo.setIconSize(QPixmap(":/down-arrow.png").scaled(22,22).size())
        self.producto_combo.setEnabled(False)
        self.producto_combo.setEditable(True)
        self.producto_combo.lineEdit().setPlaceholderText("Buscar producto...")
        self.producto_combo.setMinimumHeight(40)

        # Configurar QCompleter y QSortFilterProxyModel para el buscador
        self.producto_proxy_model = QSortFilterProxyModel(self)
        self.producto_proxy_model.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.producto_completer = QCompleter(self)
        self.producto_completer.setModel(self.producto_proxy_model)
        self.producto_completer.setCompletionMode(QCompleter.CompletionMode.UnfilteredPopupCompletion)
        self.producto_combo.setCompleter(self.producto_completer)
        self.producto_combo.lineEdit().textEdited.connect(self.on_producto_text_changed)
        self.producto_combo.setStyleSheet(COMBO_BOX_STYLE)

        form_layout.addLayout(producto_num_productos)
        form_layout.addWidget(self.producto_combo)

        info_producto_layout = QVBoxLayout()
        info_producto_layout.setAlignment(Qt.AlignCenter)
        info_producto_layout.setContentsMargins(100, 20, 100, 20)
        info_producto_layout.setSpacing(20)

        producto_label_input = QHBoxLayout()
        nombre_producto_seleccionado_label = QLabel("Producto seleccionado:", self)
        nombre_producto_seleccionado_label.setFont(fuente)
        nombre_producto_seleccionado_label.setStyleSheet(LABEL_STYLE)
        self.nombre_seleccion = QLineEdit("", self)
        self.nombre_seleccion.setFont(fuente)
        self.nombre_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
        self.nombre_seleccion.setReadOnly(True)
        self.nombre_seleccion.setCursor(Qt.IBeamCursor)
        producto_label_input.addWidget(nombre_producto_seleccionado_label)
        producto_label_input.addWidget(self.nombre_seleccion)

        categoria_codigo_celda_layout = QHBoxLayout()

        categoria_label_input = QHBoxLayout()
        categoria_producto_seleccionado_label = QLabel("Categoría:", self)
        categoria_producto_seleccionado_label.setFont(fuente)
        categoria_producto_seleccionado_label.setStyleSheet(LABEL_STYLE)
        self.categoria_seleccion = QLineEdit("", self)
        self.categoria_seleccion.setFont(fuente)
        self.categoria_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
        self.categoria_seleccion.setReadOnly(True)
        #Agregar el cursos del seleccion de texto
        self.categoria_seleccion.setCursor(Qt.IBeamCursor)
        categoria_label_input.addWidget(categoria_producto_seleccionado_label)
        categoria_label_input.addWidget(self.categoria_seleccion)
        categoria_codigo_celda_layout.addLayout(categoria_label_input)

        codigo_label_input = QHBoxLayout()
        codigo_producto_seleccionado_label = QLabel("Código:", self)
        codigo_producto_seleccionado_label.setFont(fuente)
        codigo_producto_seleccionado_label.setStyleSheet(LABEL_STYLE)
        self.codigo_seleccion = QLineEdit("", self)
        self.codigo_seleccion.setFont(fuente)
        self.codigo_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
        self.codigo_seleccion.setReadOnly(True)
        self.codigo_seleccion.setCursor(Qt.IBeamCursor)
        codigo_label_input.addWidget(codigo_producto_seleccionado_label)
        codigo_label_input.addWidget(self.codigo_seleccion)
        categoria_codigo_celda_layout.addLayout(codigo_label_input)

        celda_label_input = QHBoxLayout()
        celda_producto_seleccionado_label = QLabel("Celda:", self)
        celda_producto_seleccionado_label.setFont(fuente)
        celda_producto_seleccionado_label.setStyleSheet(LABEL_STYLE)
        self.celda_seleccion = QLineEdit("", self)
        self.celda_seleccion.setFont(fuente)
        self.celda_seleccion.setStyleSheet(INPUT_STYLE_DISABLED)
        self.celda_seleccion.setReadOnly(True)
        self.celda_seleccion.setCursor(Qt.IBeamCursor)
        celda_label_input.addWidget(celda_producto_seleccionado_label)
        celda_label_input.addWidget(self.celda_seleccion)
        categoria_codigo_celda_layout.addLayout(celda_label_input)

        info_producto_layout.addLayout(producto_label_input)
        info_producto_layout.addLayout(categoria_codigo_celda_layout)


        # Botones
        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignCenter)
        button_layout.setSpacing(30)

        cargar_button = QPushButton("Cargar archivo", self)
        cargar_button.setFont(fuente)
        cargar_button.setCursor(Qt.PointingHandCursor)
        cargar_button.setStyleSheet(BUTTON_STYLE)
        cargar_button.setIcon(QIcon.fromTheme("document-open"))
        cargar_button.clicked.connect(self.ventana_escoger_archivo)

        buscar_button = QPushButton("Buscar código", self)
        buscar_button.setFont(fuente)
        buscar_button.setCursor(Qt.PointingHandCursor)
        buscar_button.setStyleSheet(BUTTON_STYLE)
        buscar_button.setIcon(QIcon.fromTheme("edit-find"))
        buscar_button.clicked.connect(self.mostrar_info_emergente)
        
        self.button_edit = QPushButton("Editar", self)
        self.button_edit.setFont(fuente)
        self.button_edit.setCursor(Qt.PointingHandCursor)
        self.button_edit.setStyleSheet(BUTTON_STYLE)
        self.button_edit.setIcon(QIcon.fromTheme("document-edit"))
        self.button_edit.clicked.connect(self.habilitar_campos_edicion)
        
        self.button_añadir = QPushButton("Añadir Producto", self)
        self.button_añadir.setFont(fuente)
        self.button_añadir.setCursor(Qt.PointingHandCursor)
        self.button_añadir.setStyleSheet(BUTTON_STYLE)
        # Añadir icono de los predeterminados
        self.button_añadir.setIcon(QIcon.fromTheme("document-save-as"))
        self.button_añadir.clicked.connect(self.habilitar_campos_añadir)

        button_layout.addWidget(cargar_button)
        button_layout.addWidget(buscar_button)
        button_layout.addWidget(self.button_edit)
        button_layout.addWidget(self.button_añadir)

        # Ruta del archivo
        self.labelPathToFile = QLabel("", self)
        self.labelPathToFile.setFont(fuente)
        self.labelPathToFile.setAlignment(Qt.AlignCenter)
        self.labelPathToFile.setStyleSheet("font-size: 15px; font-weight: bold; color: gray;")

        self.label_mensajes_alerta = QLabel("", self)
        self.label_mensajes_alerta.setFont(fuente)
        self.label_mensajes_alerta.setAlignment(Qt.AlignCenter)
        self.label_mensajes_alerta.setWordWrap(True)
        self.label_mensajes_alerta.setStyleSheet("font-size: 15px; font-weight: bold; color: red; margin-left: 30px; margin-right: 30px;")
        self.label_mensajes_alerta.setMinimumWidth(self.width() - 200)
        self.label_mensajes_alerta.setMinimumHeight(0)

        # Agregar widgets al layout principal
        main_layout.addWidget(banner)
        main_layout.addLayout(form_layout)
        main_layout.addWidget(self.label_mensajes_alerta)
        main_layout.addLayout(info_producto_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.labelPathToFile)

        self.verificar_habilitar_campos_edicion()
        self.verificar_habilitar_campos_añadir()
        self.checkear_ruta_del_archivo()

def main():
    app = QApplication(sys.argv)
    cargar_fuentes()
    ventana_principal = VentanaPrincipal()
    ventana_principal.show()
    sys.exit(app.exec())
