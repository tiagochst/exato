
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import *
from PyQt4.QtGui import *
import sys
import ImageViewerUI
import search

class ImageViewer(QtGui.QMainWindow, ImageViewerUI.Ui_mainWindow):

    def __init__(self, parent=None):
        super(ImageViewer, self).__init__(parent)
        self.setupUi(self)
        self.initConnections()

    def about (self): 
        QMessageBox.about(self, 
                          "Exato 2013", 
                          "<b>Programa: Encontra aprovados</b>"
                          "<br> Autor: <i>Tiago Chedraoui Silva</i>"
                          "<br> Versao: 0.1 </i>"
        )

    def extensionError (self,extension): 
        QMessageBox.critical(None,"Erro de extensao","O arquivo esperado e um "+ extension)

    def howTo (self): 
        QMessageBox.about(self, 
                          "Programa de verificacao de" 
                          " aprovados vestibulares",
                          "<br>O programa recebe como entrada "
                          "um arquivo xlsx com os nomes dos alunos" 
                          "do exato, um arquivo de chamada de um "
                          "dos seguintes vestibulares:"
                          "<br> = Fuvest : Lista txt de aprovados"
                          "<br> = Unicamp: Lista txt de aprovados"
                          "<br> = Ufscar : Lista pdf de aprovados"
                          "<br> = Unesp: : Lista pdf de aprovados "
                          "<br><br>Saida: "
                          "<br>Lista de aprovados"
                      )

    # Create connections of buttons    
    def initConnections(self):      

        self.actionComo_Usar.triggered.connect(self.howTo)
        self.actionSobre.triggered.connect(self.about)
        self.fileVestibular.pressed.connect(self.selectVestibular)
        self.fileExato.pressed.connect(self.selectExato)
        self.pbFind.pressed.connect(self.compareLists)

    def compareLists(self):
        # verify inputs
        if(self.pathExato.text()!= ""):
            if(self.pathVestibular.text()!= ""):

                if(self.rbUfscar.isChecked()):
                    if str(self.pathVestibular.text()).endswith('.pdf'):
                        search.find(self.pathExato.text(),
                                    self.pathVestibular.text(),
                                    search.readUfscar)
                    else:
                        self.extensionError("pdf")
                elif(self.rbUnicamp.isChecked()):
                    if str(self.pathVestibular.text()).endswith('.txt'):
                        search.find(self.pathExato.text(),
                                    self.pathVestibular.text(),
                                    search.readUnicamp)
                    else:
                         self.extensionError("txt")
                elif(self.rbUsp.isChecked()):
                    if str(self.pathVestibular.text()).endswith('.txt'):
                        search.find(self.pathExato.text(),
                                    self.pathVestibular.text(),
                                    search.readUsp)
                    else:
                          self.extensionError("txt")
                elif(self.rbUnesp.isChecked()):
                    if str(self.pathVestibular.text()).endswith('.pdf'):
                        search.find(self.pathExato.text(),
                                    self.pathVestibular.text(),
                                    search.readUnesp)
                    else:
                          self.extensionError("pdf")

            else:
                QMessageBox.critical(None,"Erro",                                     "Selecione o arquivo de vestibular ")
        else:
            QMessageBox.critical(None,"Erro",
            "Selecione o arquivo de alunos")

    # Copy the path of the selected file to the qline
    def selectExato(self):
        fname = QtGui.QFileDialog.getOpenFileName(self, 'Arquivo de alunos xlsx')
        self.pathExato.setText(fname) 

    # Copy the path of the selected file to the qline
    def selectVestibular(self):
        fname = QtGui.QFileDialog.getOpenFileName(self, 'Lista de aprovados no vestibular')
        self.pathVestibular.setText(fname) 

    # Put the window in the center of the screen 
    def center(self):
        screen = QtGui.QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width()-size.width())/2, (screen.height()-size.height())/2)

    def main(self):
        self.center()
        self.show()

if __name__=='__main__':
    app = QtGui.QApplication(sys.argv)
    imageViewer = ImageViewer()
    imageViewer.main()
    app.exec_()

