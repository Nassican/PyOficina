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
    QListView,
    QAbstractItemView,
    QCompleter
)
from PySide6.QtCore import Qt, QSortFilterProxyModel, QFile, QResource
import PySide6
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
        self.setStyleSheet("background-color: white; color: black;")

    def mostrar_ventana_emergente(self):
        if not self.pathToFileComplete:
            self.messageBox.setText("No hay archivo seleccionado")
            self.messageBox.exec()
            return

        self.buscar_codigo_producto()
        ventana_emergente = QDialog(self)
        ventana_emergente.setWindowTitle("Ventana Emergente")
        ventana_emergente.resize(800, 450)
        ventana_emergente.setStyleSheet("background-color: white; color: black;")

        layout = QVBoxLayout(ventana_emergente)
        layout.setAlignment(Qt.AlignCenter)

        fuente = QFont("Red Hat Text", 20)
        textNomProd = QLabel(f"Producto: {self.varTextoNomProd.upper()}", ventana_emergente)
        textNomProd.setFont(fuente)
        textNomProd.setAlignment(Qt.AlignCenter)
        textNomProd.setStyleSheet("font-size: 50px; font-weight: bold;")
        textNomProd.setWordWrap(True)

        textInfoProd = QLabel("Este producto pertenece a la categoría:", ventana_emergente)
        textInfoProd.setFont(fuente)

        textTipoProd = QLabel(self.varTextTipoProd, ventana_emergente)
        textTipoProd.setFont(fuente)
        textTipoProd.setStyleSheet("font-size: 30px; font-weight: bold;")

        InfoTipoProd = QHBoxLayout()
        InfoTipoProd.setAlignment(Qt.AlignCenter)
        InfoTipoProd.addWidget(textInfoProd)
        InfoTipoProd.addWidget(textTipoProd)

        textCodLabel = QLabel("El código del producto es:", ventana_emergente)
        textCodLabel.setFont(fuente)
        textCodLabel.setAlignment(Qt.AlignCenter)
        textCodLabel.setWordWrap(True)

        numCodProd = QLabel(f"{self.varNumCodProd}", ventana_emergente)
        numCodProd.setFont(fuente)
        numCodProd.setAlignment(Qt.AlignCenter)

        buttonCerrar = QPushButton("Cerrar", ventana_emergente)
        buttonCerrar.setFont(fuente)
        buttonCerrar.setCursor(Qt.PointingHandCursor)
        buttonCerrar.setStyleSheet(BUTTON_STYLE)
        buttonCerrar.clicked.connect(ventana_emergente.close)
        buttonCerrar.setMaximumSize(200, 50)
        buttonCerrar.setMinimumSize(200, 50)
        
        layoutBoton = QVBoxLayout()
        layoutBoton.setAlignment(Qt.AlignCenter | Qt.AlignTop)
        layoutBoton.addWidget(buttonCerrar)
        layoutBoton.setContentsMargins(0, 50, 0, 0)

        layout.addWidget(textNomProd)
        layout.addLayout(InfoTipoProd)
        layout.addWidget(textCodLabel)
        layout.addWidget(numCodProd)
        layout.addLayout(layoutBoton)

        #ventana_emergente.exec()

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

    def on_producto_text_changed(self, text):
        self.producto_proxy_model.setFilterFixedString(text)

    def buscar_codigo_producto(self):
        categoria_texto = self.categoria_combo.currentText()
        producto_texto = self.producto_combo.currentText()
        codigo, columna_letra, fila_numero = excel.buscar_codigo_producto(categoria_texto, producto_texto, self.pathToFileComplete)
        celda_codigo = f"{columna_letra}{fila_numero}"
        self.varTextoNomProd = producto_texto
        self.varTextTipoProd = categoria_texto
        self.varNumCodProd = f"{codigo} (celda {celda_codigo})"
        
        self.nombre_seleccion.setText(self.varTextoNomProd)
        self.categoria_seleccion.setText(self.varTextTipoProd)
        self.codigo_seleccion.setText(str(codigo))
        self.celda_seleccion.setText(celda_codigo)

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
                self.cargar_categorias()
                self.onChangedCategoria()
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
            # self.nombre_input.setEnabled(True)
            # self.nombre_input.setStyleSheet(INPUT_STYLE)
            # self.nombre_producto_input.setEnabled(True)
            # self.nombre_producto_input.setStyleSheet(INPUT_STYLE)
            print(f"Archivo seleccionado, RUTA: {pathToFileComplete}")
        else:
            # Desactivar inputs y combo
            self.categoria_combo.setEnabled(False)
            self.categoria_combo.setStyleSheet(COMBO_BOX_STYLE_DISABLED)
            self.producto_combo.setEnabled(False)
            self.producto_combo.setStyleSheet(COMBO_BOX_STYLE_DISABLED)
            # self.nombre_input.setEnabled(False)
            # self.nombre_input.setStyleSheet(INPUT_STYLE_DISABLED)
            # self.nombre_producto_input.setEnabled(False)
            # self.nombre_producto_input.setStyleSheet(INPUT_STYLE_DISABLED)
            print(f"No hay archivo seleccionado, RUTA: {pathToFileComplete}")

    def initUI(self):
        self.setWindowTitle("Oficina de Planeación")
        self.resize(900, 525)
        self.setMaximumSize(900, 525)
        self.setMinimumSize(900, 525)

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

        # Input de nombre
        # nombre_label = QLabel("Nombre:", self)
        # nombre_label.setFont(fuente)
        # nombre_label.setStyleSheet(LABEL_STYLE)
        # self.nombre_input = QLineEdit(self)
        # self.nombre_input.setFont(fuente)
        # self.nombre_input.setStyleSheet(INPUT_STYLE)
        # self.nombre_input.setEnabled(False)
        # self.nombre_input.setPlaceholderText("Nombre Producto")

        #form_layout.addWidget(nombre_label)
        #form_layout.addWidget(self.nombre_input)

        # Input de nombre de producto
        #nombre_producto_label = QLabel("Nombre del producto:", self)
        #nombre_producto_label.setFont(fuente)
        #nombre_producto_label.setStyleSheet(LABEL_STYLE)
        #self.nombre_producto_input = QLineEdit(self)
        #self.nombre_producto_input.setFont(fuente)
        #self.nombre_producto_input.setStyleSheet(INPUT_STYLE)
        #self.nombre_producto_input.setEnabled(False)
        #self.nombre_producto_input.setPlaceholderText("Nombre del producto")

        #form_layout.addWidget(nombre_producto_label)
        #form_layout.addWidget(self.nombre_producto_input)

        # Nuevo layout para informacion del producto (Nombre, categoria, codigo y celda)

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
        self.nombre_seleccion.setStyleSheet(INPUT_STYLE)
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
        self.categoria_seleccion.setStyleSheet(INPUT_STYLE)
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
        self.codigo_seleccion.setStyleSheet(INPUT_STYLE)
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
        self.celda_seleccion.setStyleSheet(INPUT_STYLE)
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

        cargar_button = QPushButton("Cargar archivo", self)
        cargar_button.setFont(fuente)
        cargar_button.setCursor(Qt.PointingHandCursor)
        cargar_button.setStyleSheet(BUTTON_STYLE)
        cargar_button.clicked.connect(self.ventana_escoger_archivo)

        buscar_button = QPushButton("Buscar código", self)
        buscar_button.setFont(fuente)
        buscar_button.setCursor(Qt.PointingHandCursor)
        buscar_button.setStyleSheet(BUTTON_STYLE)
        buscar_button.clicked.connect(self.mostrar_ventana_emergente)

        button_layout.addWidget(cargar_button)
        button_layout.addWidget(buscar_button)

        # Ruta del archivo
        self.labelPathToFile = QLabel("", self)
        self.labelPathToFile.setFont(fuente)
        self.labelPathToFile.setAlignment(Qt.AlignCenter)
        self.labelPathToFile.setStyleSheet("font-size: 15px; font-weight: bold; color: gray;")

        # Agregar widgets al layout principal
        main_layout.addWidget(banner)
        main_layout.addLayout(form_layout)
        main_layout.addLayout(info_producto_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.labelPathToFile)

        self.checkear_ruta_del_archivo()

def main():
    app = QApplication(sys.argv)
    cargar_fuentes()
    ventana_principal = VentanaPrincipal()
    ventana_principal.show()
    sys.exit(app.exec())
