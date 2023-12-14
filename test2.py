import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QFrame, QVBoxLayout, QSizePolicy
from PySide6.QtUiTools import QUiLoader
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np
class MyMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Cargar el archivo de interfaz de usuario
        loader = QUiLoader()
        self.ui = loader.load("mainwindow.ui")
        self.setCentralWidget(self.ui)

        # Acceder al QFrame desde el archivo de interfaz de usuario
        frame = self.ui.findChild(QFrame, "graph")

        # Crear una figura de Matplotlib
        fig = Figure(figsize=(5, 4), dpi=100)
        canvas = FigureCanvas(fig)
        ax = fig.add_subplot(111)

        # Crear datos para el gráfico
        x = np.linspace(0, 5, 100)
        y = np.sin(x)

        # Dibujar en la figura
        ax.plot(x, y)
        ax.set_title('Gráfico de Matplotlib')

        # Crear un layout vertical para el QFrame y agregar el lienzo de Matplotlib
        layout = QVBoxLayout(frame)
        layout.addWidget(canvas)

        # Establecer el layout del QFrame
        frame.setLayout(layout)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MyMainWindow()
    mainWindow.show()
    sys.exit(app.exec_())