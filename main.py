# import sys
# from PyQt5.QtWidgets import QApplication, QTextEdit, QLabel, QPushButton, QFileDialog, QMainWindow
# from PyQt5.QtWidgets import QWidget, QTextEdit, QVBoxLayout, QInputDialog
# from PyQt5 import QtCore
#
#
# def directory_button_func(self):  # Choose button
#     directory_button = QPushButton()
#     directory_button.setText("Выберите папку")
#     directory_button.setMinimumSize(QtCore.QSize(70, 40))
#     directory_button.clicked.connect(self.input_dialog)
#     self.vbox.addWidget(directory_button)
#
#
# def path_user_input(self):  # Show directory chosen by user
#     path = QTextEdit()
#     path.resize(40, 40)
#     path.setMinimumSize(QtCore.QSize(40, 40))
#     self.path_exel = path
#     self.vbox.addWidget(self.path_exel)
#
#
# class NewWindow(QMainWindow):  # Create window
#     def __init__(self):
#         QMainWindow.__init__(self)
#         self.resize(500, 500)
#         self.setWindowTitle("RevitHelper")
#         self.centralWidget = QWidget(self)
#         self.centralWidget.resize(500, 500)
#         self.vbox = QVBoxLayout(self.centralWidget)
#         directory_button_func(self)
#         path_user_input(self)
#
#     def input_dialog(self):  # Returns user input
#         val_return = QFileDialog.getExistingDirectory()
#         text_to_print = str(val_return)
#         self.path.setText(text_to_print)
#
#
# def main():
#     app = QApplication(sys.argv)
#     win = NewWindow()
#     win.show()
#     app.exec_()
#
#
# # if __name__ == '__main__':
# #     main()
#
#
import sys
from PyQt5.QtWidgets import QWidget, QPushButton, QApplication
from PyQt5.QtCore import QCoreApplication


class Example(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        qbtn = QPushButton('Quit', self)
        qbtn.clicked.connect(QCoreApplication.instance().quit)
        qbtn.resize(qbtn.sizeHint())
        qbtn.move(50, 50)

        self.setGeometry(300, 300, 250, 150)
        self.setWindowTitle('Quit button')
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
