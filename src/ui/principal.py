import sys
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QLineEdit, QVBoxLayout, QPushButton, QDialog, QHBoxLayout, QFileDialog, QMessageBox
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontDatabase, QIcon
import os
from ..back import funciones as excel

# Ruta de este archivo
path = os.path.dirname(os.path.abspath(__file__))
window_icon_path = os.path.join(path, "img/escudo.png")
GREEN_COLOR = "#00FF00"


def cargar_fuentes():
    # Cargar la fuente Frostbite
    QFontDatabase.addApplicationFont(
        os.path.join(path, "fonts/RedHatText-Bold.ttf"))
    QFontDatabase.addApplicationFont(
        os.path.join(path, "fonts/RedHatText-Regular.ttf"))
    QFontDatabase.addApplicationFont(
        os.path.join(path, "fonts/Overpass-Bold.ttf"))

    familyArray = QFontDatabase().families()
    # print(familyArray)


class MessageWidget(QMessageBox):
    def __init__(self):
        super(MessageWidget, self).__init__()
        self.setWindowTitle("Oficina de Planeación")
        self.button_message = QPushButton("Aceptar")
        self.button_message.setCursor(Qt.PointingHandCursor)
        self.setWindowIcon(QIcon(window_icon_path))
        self.addButton(self.button_message, QMessageBox.AcceptRole)
        self.setIcon(QMessageBox.Critical)
        self.setStyleSheet("""
            QMessageBox{
            border: 10px;
            font-size: 20px;
            }
            QPushButton {
            font-size: 16px;
            background-color: green;
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
                           )


class VentanaPrincipal(QWidget):

    pathToFileComplete = ""

    def __init__(self):
        super().__init__()
        self.initUI()
        self.setWindowIcon(QIcon(window_icon_path))
        self.messageBox = MessageWidget()
        self.setStyleSheet("background-color: white; color: black;")

    def mostrar_ventana_emergente(self):
        # Crear una ventana emergente (QDialog)
        ventana_emergente = QDialog(self)
        ventana_emergente.setWindowTitle("Ventana Emergente")
        ventana_emergente.resize(800, 450)
        ventana_emergente.setMaximumHeight(450)
        ventana_emergente.setMinimumHeight(450)
        ventana_emergente.setMaximumWidth(800)
        ventana_emergente.setMinimumWidth(800)
        fuente = QFont("Red Hat Text", 20)
        self.setStyleSheet("background-color: white; color: black;")

        # Agregar un layout vertical a la ventana emergente
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)

        labelVacio = QWidget()
        labelVacio.setMaximumHeight(50)
        labelVacio.setMinimumHeight(50)

        self.varTextoNomProd = "Nombre del Producto"
        self.varTextTipoProd = "Nombre Categoría"
        self.varNumCodProd = "00000000"
        if self.pathToFileComplete == "":
            self.messageBox.setText("No hay archivo seleccionado")
            self.messageBox.exec()
            return
        else:
            self.buscar_codigo_producto(self.pathToFileComplete)
            # Pasar a mayusculas
            self.varTextoNomProd = self.varTextoNomProd.upper()

        textNomProd = QLabel(f"Producto: {self.varTextoNomProd}")
        textNomProd.setFont(fuente)
        textNomProd.setAlignment(Qt.AlignCenter)
        textNomProd.setStyleSheet("font-size: 50px; font-weight: bold;")

        InfoTipoProd = QHBoxLayout()
        InfoTipoProd.setAlignment(Qt.AlignCenter)
        textInfoProd = QLabel("Este producto pertenece a la categoría:")
        textInfoProd.setFont(fuente)

        textTipoProd = QLabel(self.varTextTipoProd)
        textTipoProd.setStyleSheet("font-size: 30px; font-weight: bold;")
        textTipoProd.setFont(fuente)

        InfoTipoProd.addWidget(textInfoProd)
        InfoTipoProd.addWidget(textTipoProd)

        textCodLabel = QLabel("El código del producto es:")
        textCodLabel.setFont(fuente)
        textCodLabel.setAlignment(Qt.AlignCenter)
        numCodProd = QLabel(f"{self.varNumCodProd}")
        numCodProd.setFont(fuente)
        numCodProd.setAlignment(Qt.AlignCenter)

        buttonCerrar = QPushButton("Cerrar")
        buttonCerrar.setFont(fuente)
        buttonCerrar.setCursor(Qt.PointingHandCursor)
        style_boton = """
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
        buttonCerrar.setStyleSheet(style_boton)
        buttonCerrar.setMaximumWidth(100)
        buttonCerrar.setMinimumWidth(100)
        buttonCerrar.setMinimumHeight(40)
        buttonCerrar.setMaximumHeight(40)
        buttonCerrar.clicked.connect(ventana_emergente.close)
        # Widget Boton
        layoutBoton = QVBoxLayout()
        layoutBoton.setAlignment(Qt.AlignCenter | Qt.AlignTop)
        layoutBoton.addWidget(buttonCerrar)

        layout.addWidget(textNomProd)
        layout.addLayout(InfoTipoProd)
        layout.addWidget(labelVacio)
        layout.addWidget(textCodLabel)
        layout.addWidget(numCodProd)
        layout.addWidget(labelVacio)
        layout.addLayout(layoutBoton)

        ventana_emergente.setLayout(layout)

        # Mostrar la ventana emergente
        ventana_emergente.exec()

    def cargar_categorias(self, rutaExcel):
        categorias = excel.listar_categorias(rutaExcel)
        self.categoria_combo.clear()
        self.categoria_combo.addItems(categorias)

    def onChangedCategoria(self, rutaExcel):
        categoria = self.categoria_combo.currentText()
        productos = excel.listar_productos(categoria, rutaExcel)
        self.producto_combo.clear()
        self.producto_combo.addItems(productos)

    def buscar_codigo_producto(self, archivo_excel):
        categoria_texto = self.categoria_combo.currentText()
        producto_texto = self.producto_combo.currentText()
        codigo = excel.buscar_codigo_producto(
            categoria_texto, producto_texto, archivo_excel)

        print(f"Producto: {producto_texto}, Categoria: {
              categoria_texto}, Codigo: {codigo}")

        self.varTextoNomProd = producto_texto
        self.varTextTipoProd = categoria_texto
        self.varNumCodProd = codigo

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
                    pathToFileLabel = "..." + \
                        pathToFileLabel[-maximosCaracteres:]

                self.pathToFileComplete = pathToFileComplete
                self.labelPathToFile.setText(pathToFileLabel)
                self.checkear_ruta_del_archivo()
                self.cargar_categorias(pathToFileComplete)
                self.onChangedCategoria(pathToFileComplete)
                print(pathToFileComplete)

    def checkear_ruta_del_archivo(self):
        # Verificar si hay un archivo seleccionado
        pathToFileComplete = self.pathToFileComplete

        if pathToFileComplete != "":
            # Activar inputs y combo
            self.categoria_combo.setEnabled(True)
            self.categoria_combo.setStyleSheet(self.style_combo_box)
            self.producto_combo.setEnabled(True)
            self.producto_combo.setStyleSheet(self.style_combo_box)
            self.nombre_input.setEnabled(True)
            self.nombre_input.setStyleSheet(self.style_input)
            self.nombre_producto_input.setEnabled(True)
            self.nombre_producto_input.setStyleSheet(self.style_input)
            print(f"Archivo seleccionado, RUTA: {pathToFileComplete}")
        else:
            # Desactivar inputs y combo
            self.categoria_combo.setEnabled(False)
            self.categoria_combo.setStyleSheet(self.style_combo_box_disabled)
            self.producto_combo.setEnabled(False)
            self.producto_combo.setStyleSheet(self.style_combo_box_disabled)
            self.nombre_input.setEnabled(False)
            self.nombre_input.setStyleSheet(self.style_input_disabled)
            self.nombre_producto_input.setEnabled(False)
            self.nombre_producto_input.setStyleSheet(self.style_input_disabled)
            print(f"No hay archivo seleccionado, RUTA: {pathToFileComplete}")

    def initUI(self):
        # Configuración de la ventana
        self.setWindowTitle("Oficina de Planeacion")
        self.resize(900, 525)
        self.setMaximumHeight(525)
        self.setMinimumHeight(525)
        self.setMaximumWidth(900)
        self.setMinimumWidth(900)
        fuente = QFont("Red Hat Text", 20)

        # Banner superior verde
        banner = QLabel()
        banner.setMaximumHeight(100)
        banner.setMinimumWidth(445)
        banner.setText("Oficina de Planeación")
        banner.setFont(fuente)
        # Estilo responsivo del banner
        banner_style = """
            QLabel {
            background-color: green;
            color: white;
            font-size: 40px;
            font-weight: bold;
            padding: 10px;
            }
        """
        banner.setStyleSheet(banner_style)

        style_combo_box = """
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
            padding: 5px;
            border-left-width: 20px;
            border-left-color: darkgray;
            }
            QComboBox QAbstractItemView {
            border: 1px solid black;
            border-radius: 1px;
            padding: 5px;
            }
            QComboBox:item:selected {
                padding-left: 10px;  /* move text right to make room for tick mark */
                border: 2px solid black;
                border-radius: 10px;
            }
        """

        style_combo_box_disabled = """
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

        style_input_disabled = """
            QLineEdit {
            font-size: 20px;
            background-color: lightgray;
            color: gray;
            border: 1px solid gray;
            border-radius: 10px;
            padding: 5px;
            }
        """

        style_input = """
            QLineEdit {
            font-size: 20px;
            background-color: white;
            color: black;
            border: 1px solid gray;
            border-radius: 10px;
            padding: 5px;
            }
        """
        self.style_input = style_input
        self.style_combo_box = style_combo_box
        self.style_combo_box_disabled = style_combo_box_disabled
        self.style_input_disabled = style_input_disabled

        # Categoría (ComboBox)
        widgetLayout = QVBoxLayout()
        widgetLayout.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)

        widgetForm = QWidget()
        widgetForm.setMaximumWidth(500)
        widgetForm.setMinimumWidth(350)
        widgetForm.setMaximumHeight(500)
        widgetForm.setMinimumHeight(350)
        widgetForm.setStyleSheet("border-radius: 10px;")

        layoutInput = QVBoxLayout()
        label_style = "font-size: 20px; font-weight: bold;"

        categoria_label = QLabel("Categoría:")
        categoria_label.setFont(fuente)
        categoria_label.setStyleSheet(label_style)

        self.categoria_combo = QComboBox()
        self.categoria_combo.setMaximumHeight(40)
        self.categoria_combo.setMinimumHeight(40)
        self.categoria_combo.setFont(fuente)
        self.categoria_combo.setStyleSheet(style_combo_box)
        # Agregar opciones a la lista desplegable de Categoría
        self.categoria_combo.addItem("Seleccione una categoría")
        # ... (Agregar más opciones según sea necesario)
        self.categoria_combo.currentIndexChanged.connect(
            lambda: self.onChangedCategoria(self.pathToFileComplete))

        # Producto (ComboBox)
        producto_label = QLabel("Producto:")
        producto_label.setFont(fuente)
        producto_label.setStyleSheet(label_style)
        self.producto_combo = QComboBox()
        self.producto_combo.setMaximumHeight(40)
        self.producto_combo.setMinimumHeight(40)
        self.producto_combo.setFont(fuente)

        self.producto_combo.setStyleSheet(style_combo_box)
        # Agregar opciones a la lista desplegable de Producto
        self.producto_combo.addItem("Seleccione un producto")
        # ... (Agregar más opciones según sea necesario)

        # Input para nombre
        nombre_label = QLabel("Ingrese Nombre:")
        nombre_label.setFont(fuente)
        nombre_label.setStyleSheet(label_style)
        self.nombre_input = QLineEdit()
        self.nombre_input.setFont(fuente)
        self.nombre_input.setPlaceholderText("Nombre Producto")
        self.nombre_input.setStyleSheet(style_input)

        # Input para nombre de producto
        nombre_producto_label = QLabel("Producto:")
        nombre_producto_label.setFont(fuente)
        nombre_producto_label.setStyleSheet(label_style)
        self.nombre_producto_input = QLineEdit()
        self.nombre_producto_input.setFont(fuente)
        self.nombre_producto_input.setPlaceholderText("Nombre del producto")
        self.nombre_producto_input.setStyleSheet(style_input)

        # Widget Boton
        layoutBoton = QHBoxLayout()
        layoutBoton.setAlignment(Qt.AlignCenter | Qt.AlignTop)

        # Boton Enviar
        boton = QPushButton("Enviar")
        boton.setFont(fuente)
        boton.setCursor(Qt.PointingHandCursor)

        botonArchivo = QPushButton("Archivo")
        botonArchivo.setFont(fuente)
        botonArchivo.setCursor(Qt.PointingHandCursor)

        # Estilo del botón con limite de tamaño
        style_boton = """
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

        self.style_boton = style_boton

        boton.setStyleSheet(style_boton)
        boton.setMaximumWidth(100)
        boton.setMinimumWidth(100)
        boton.setMinimumHeight(40)
        boton.setMaximumHeight(40)

        botonArchivo.setStyleSheet(style_boton)
        botonArchivo.setMaximumWidth(100)
        botonArchivo.setMinimumWidth(100)
        botonArchivo.setMinimumHeight(40)
        botonArchivo.setMaximumHeight(40)

        labelPathToFile = QLabel("")
        labelPathToFile.setFont(fuente)
        labelPathToFile.setStyleSheet(
            "font-size: 15px; font-weight: bold; color: gray;")
        labelPathToFile.setWordWrap(True)
        labelPathToFile.setAlignment(Qt.AlignCenter)
        self.labelPathToFile = labelPathToFile

        layoutBoton.addWidget(boton)
        layoutBoton.addWidget(botonArchivo)
        # layoutBoton.addWidget(labelPathToFile)

        # Conectar boton enviar con la ventana emergente y que no se cierre
        boton.clicked.connect(self.mostrar_ventana_emergente)
        botonArchivo.clicked.connect(self.ventana_escoger_archivo)
        layoutInput.addWidget(categoria_label)
        layoutInput.addWidget(self.categoria_combo)
        layoutInput.addWidget(producto_label)
        layoutInput.addWidget(self.producto_combo)
        layoutInput.addWidget(nombre_label)
        layoutInput.addWidget(self.nombre_input)
        layoutInput.addWidget(nombre_producto_label)
        layoutInput.addWidget(self.nombre_producto_input)

        # widgetLayout = QWidget()
        # widgetForm = QWidget()
        # layoutInput = QVBoxLayout()

        widgetForm.setLayout(layoutInput)
        widgetLayout.addWidget(widgetForm)

        # Diseño de la interfaz
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(banner)
        layout.addLayout(widgetLayout)
        layout.addLayout(layoutBoton)
        layout.addWidget(labelPathToFile)

        self.checkear_ruta_del_archivo()
        self.setLayout(layout)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    cargar_fuentes()
    window = VentanaPrincipal()
    window.show()
    sys.exit(app.exec())
