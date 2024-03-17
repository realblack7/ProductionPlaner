import datetime
import sys
import os
import shutil
import re
from PIL import Image, ImageQt
import zipfile
from openpyxl import Workbook, load_workbook
import configparser
from io import StringIO, BytesIO
from lxml import html

from PyQt6.QtCore import Qt, QObject, QThread, pyqtSignal, QRegularExpression
from PyQt6.QtWidgets import (
    QApplication,
    QLabel,
    QMainWindow,
    QPushButton,
    QWidget,
    QGridLayout,
    QLineEdit,
    QHBoxLayout,
    QVBoxLayout,
    QListWidget,    
    QSpinBox,    
    QTextEdit,
    QProgressBar,    
    QComboBox,
    QAbstractItemView,
    QTreeWidget,
    QTreeWidgetItem,
    QListWidgetItem,
    QFileDialog,
    QMessageBox,
    QDateEdit,
    QRadioButton,
    QCheckBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QCompleter, 
    QTabWidget   
    
                
)
from PyQt6.QtGui import QIcon, QPixmap, QAction, QActionGroup, QMovie, QValidator, QDoubleValidator, QIntValidator, QRegularExpressionValidator

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'realblack7.productionmanager.v.1'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class AddBatchWindow(QWidget):
    added = pyqtSignal(list)
    finished = pyqtSignal()

    def __init__(self, attr1, attr2, attrPack, attrLab):
        super().__init__()      

        self.attr1 = attr1
        self.attr2 = attr2        
        self.attrPackaging = attrPack
        self.attrLab = attrLab        

        self.closeMenu = True
        self.setWindowTitle('Charge hinzufügen')
        self._createGUI()

    def _createGUI(self):        
        layout2 = QVBoxLayout()
        layout3 = QVBoxLayout()
        layout4 = QHBoxLayout() 
        layout5 = QVBoxLayout()   
        layout6 = QHBoxLayout()   
        
        self.listCostumer = QComboBox()  
        self.listCostumer.addItem('')                 
        self.listCostumer.addItems(self.attr1)        
        self.listCostumer.setEditable(True) 
        self.listCostumer.InsertPolicy.InsertAlphabetically        
        self.listCostumer.setFixedWidth(150)

        rx = QRegularExpression("SP\\d{1,9}")
        self.listDispo = QLineEdit() 
        self.listDispo.setFixedWidth(100) 
        self.listDispo.setMaxLength(8)
        self.listDispo.setText('SP')
        self.listDispo.setValidator(QRegularExpressionValidator(rx, self))

        rx2 = QRegularExpression("32.\\d{1,4}") 
        self.listArticle = QComboBox()  
        self.listArticle.addItem('32.')                 
        self.listArticle.addItems(self.attr2)        
        self.listArticle.setEditable(True) 
        self.listArticle.InsertPolicy.InsertAlphabetically        
        self.listArticle.setValidator(QRegularExpressionValidator(rx2, self))        
        

        self.listDeliveryDate = QDateEdit()
        self.listDeliveryDate.setFixedWidth(100)
        self.listDeliveryDate.setDate(datetime.datetime.now() + datetime.timedelta(days=2))
        self.listDeliveryDate.setMouseTracking(False)
        
        self.listBatchSize = QLineEdit()
        self.listBatchSize.setFixedWidth(100)        
        self.listBatchSize.setValidator(QIntValidator(00, 99))

        self.labelCharge = QLabel('Kunde')
        self.labelDispo = QLabel('Dispo-Nr.')
        self.labelArticle = QLabel('Artikle-Nr.')
        self.labelDeliveryDate = QLabel('Lieferdatum')
        self.labelBatchSize = QLabel('Tonnage')   
        self.labelPackaging = QLabel('Zusatz')
        self.labelLab = QLabel('Labor')       

        self.listPackaging = QComboBox()
        self.listPackaging.addItems(self.attrPackaging)
                       
        self.listLab = QComboBox()
        self.listLab.addItems(self.attrLab)

        self.closeButton = QPushButton('Schließen')
        self.closeButton.setFixedWidth(80)
        self.closeButton.clicked.connect(self.close)  

        self.addButton = QPushButton('Hinzufügen')
        self.addButton.setFixedWidth(80)
        self.addButton.clicked.connect(self.addBatchToList)     

        layout2.addWidget(self.listCostumer)
        layout2.addWidget(self.listDispo)
        layout2.addWidget(self.listArticle)
        layout2.addWidget(self.listDeliveryDate)
        layout2.addWidget(self.listBatchSize)
        layout2.addWidget(self.listPackaging)
        layout2.addWidget(self.listLab)            

        layout3.addWidget(self.labelCharge) 
        layout3.addWidget(self.labelDispo)
        layout3.addWidget(self.labelArticle)
        layout3.addWidget(self.labelDeliveryDate)
        layout3.addWidget(self.labelBatchSize) 
        layout3.addWidget(self.labelPackaging)
        layout3.addWidget(self.labelLab)     
      
        layout4.addLayout(layout3)
        layout4.addLayout(layout2) 

        layout6.addWidget(self.addButton)
        layout6.addWidget(self.closeButton)
        layout6.addStretch()

        layout5.addLayout(layout4)
        layout5.addLayout(layout6)     

        self.setLayout(layout5)        

    def addBatchToList(self):    
                 
        batchArray = ['', '', '', '', self.listArticle.currentIndex(), '', self.listDispo.text(), self.listCostumer.currentIndex(), self.listPackaging.currentIndex(), self.listLab.currentIndex(), self.listDeliveryDate.date().toString('dd.MM.yyyy'), self.listBatchSize.text(), '' ] 

        self.added.emit(batchArray)     
                   
    def closeEvent(self, event):
        
        if self.closeMenu == True:
            event.accept()
            self.finished.emit()
        else:
            event.ignore()

class SettingshWindow(QWidget):
    added = pyqtSignal(list)
    finished = pyqtSignal()

    def __init__(self, sortBy, timenormal, timedensity, timemechanics, timereach):
        super().__init__()   

        self.closeMenu = True
        self.sortBy = sortBy
        self.timeNormal = timenormal
        self.timeDensity = timedensity
        self.timeMechanics = timemechanics
        self.timeReach = timereach
        self.setWindowTitle('Einstellungen')
        self._createGUI()

    def _createGUI(self):        
        layout2 = QVBoxLayout()
        layout3 = QVBoxLayout()
        layout4 = QHBoxLayout() 
        layout5 = QVBoxLayout() 
        layout6 = QHBoxLayout()

        
        self.sortByBox = QComboBox()                 
        self.sortByBox.addItems(['Produktionsbeginn', 'Produktionsende', 'Abholung'])        
        self.sortByBox.setEditable(True) 
        self.sortByBox.setCurrentIndex(self.sortBy)
        self.sortByBox.InsertPolicy.InsertAlphabetically        
        self.sortByBox.setFixedWidth(140)  

        rx = QRegularExpression("\\d{1,2}")
        self.timenormalLine = QLineEdit() 
        self.timenormalLine.setFixedWidth(50) 
        self.timenormalLine.setMaxLength(2)
        self.timenormalLine.setText(str(self.timeNormal))
        self.timenormalLine.setValidator(QRegularExpressionValidator(rx, self))      

        rx = QRegularExpression("\\d{1,2}")
        self.timedensityLine = QLineEdit() 
        self.timedensityLine.setFixedWidth(50) 
        self.timedensityLine.setMaxLength(2)
        self.timedensityLine.setText(str(self.timeDensity))
        self.timedensityLine.setValidator(QRegularExpressionValidator(rx, self)) 

        rx = QRegularExpression("\\d{1,2}")
        self.timemechanicsLine = QLineEdit() 
        self.timemechanicsLine.setFixedWidth(50) 
        self.timemechanicsLine.setMaxLength(2)
        self.timemechanicsLine.setText(str(self.timeMechanics))
        self.timemechanicsLine.setValidator(QRegularExpressionValidator(rx, self)) 

        rx = QRegularExpression("\\d{1,2}")
        self.timereachLine = QLineEdit() 
        self.timereachLine.setFixedWidth(50) 
        self.timereachLine.setMaxLength(2)
        self.timereachLine.setText(str(self.timeReach))
        self.timereachLine.setValidator(QRegularExpressionValidator(rx, self))    

        self.labelSort = QLabel('Sortieren nach ')   
        self.labelNormal = QLabel('Vorlauf Produktion ')
        self.labelDenisty = QLabel('Vorlauf Dichte-Messung ')
        self.labelMechanics = QLabel('Vorlauf Mechanik-Messung ')
        self.labelReach = QLabel('Vorlauf REACh-Messung ')

        self.closeButton = QPushButton('Schließen')
        self.closeButton.setFixedWidth(80)
        self.closeButton.clicked.connect(self.close)              

        self.addButton = QPushButton('Speichern')
        self.addButton.setFixedWidth(80)
        self.addButton.clicked.connect(self.saveSettings)     

        layout2.addWidget(self.sortByBox) 
        layout2.addWidget(self.timenormalLine)
        layout2.addWidget(self.timedensityLine)
        layout2.addWidget(self.timemechanicsLine)
        layout2.addWidget(self.timereachLine)      

        layout3.addWidget(self.labelSort)  
        layout3.addWidget(self.labelNormal)
        layout3.addWidget(self.labelDenisty)
        layout3.addWidget(self.labelMechanics)
        layout3.addWidget(self.labelReach)  
      
        layout4.addLayout(layout3)
        layout4.addLayout(layout2) 

        layout6.addWidget(self.addButton)
        layout6.addWidget(self.closeButton)
        layout6.addStretch()

        layout5.addLayout(layout4)
        layout5.addLayout(layout6)
              

        self.setLayout(layout5)        

    def saveSettings(self):    
                 
        settingsToSave = [self.sortByBox.currentIndex(), self.timenormalLine.text(), self.timedensityLine.text( ), self.timemechanicsLine.text(), self.timereachLine.text()] 

        self.added.emit(settingsToSave)     
                   
    def closeEvent(self, event):
        
        if self.closeMenu == True:
            event.accept()
            self.finished.emit()
        else:
            event.ignore()

class MainWindow(QMainWindow):           

    def __init__(self):
        super().__init__()
        self.w = None
        self.closeMenu = True  
        self.workingOnShiftPlan = False
        self.setLoadedFile = False      
        self.imagePath = os.path.dirname(__file__)

        self.config = configparser.ConfigParser()
        self.config.read(os.path.join(self.imagePath, 'settings.ini'))

        self.saveFilePath = self.config['PATH']['lastsaved']   
        self.sortBy = int(self.config['SETTINGS']['sortby'])  

        self.timeNormal = int(self.config['SETTINGS']['timenormal']) 
        self.timeDensity = int(self.config['SETTINGS']['timedensity']) 
        self.timeMechanics = int(self.config['SETTINGS']['timemechanics']) 
        self.timeReach = int(self.config['SETTINGS']['timereach']) 

        match self.sortBy:
            case 0:
                self.sortByColumn = 2
            case 1:
                self.sortByColumn = 3
            case 2:
                self.sortByColumn = 10

        
        

        with open(os.path.join(self.imagePath, 'data', 'customers.txt'), 'r', -1, 'utf-8') as file:
            self.attr1 = [line.rstrip() for line in file]

        with open(os.path.join(self.imagePath, 'data', 'articles.txt'), 'r', -1, 'utf-8') as file:
            self.attr2 = [line.rstrip() for line in file]  

        self.attr1.sort()  
        self.attr2.sort()   

        self.attrShift = ['F-S','S-N','N-F','N-W-S', 'W-S-N', 'F', 'S', 'N']
        self.attrPack = ['Bigbag','Oktabin','Silo','Homogenisierung']
        self.attrLab = ['-','Dichte','Mechanik','REACh']           

        self.setWindowTitle("Produktionsplaner")        
        self.setAcceptDrops(False)
        self._createMenu() 
        self._createTabs()     
        self._createPlanerViewExtruder1()   
        self._createPlanerViewExtruder2()  
        self._createPlanerViewHomogenisation()     
        self._createPlanerViewSilo() 
        self._createMaster()    
        
    def _createMenu(self):
        self.menubar = self.addToolBar('Menü')              

        self.openFileDialog = QAction(QIcon(os.path.join(self.imagePath, 'assets', 'folder-open-solid.svg')), 'Laden (Strg + O)', self)        
        self.openFileDialog.triggered.connect(self.loadFile)
        self.openFileDialog.setShortcut("Ctrl+O")

        self.saveFile = QAction(QIcon(os.path.join(self.imagePath, 'assets', 'floppy-disk-solid.svg')), 'Speichern  (Strg + S)', self)        
        self.saveFile.triggered.connect(self.performSaveFile)
        self.saveFile.setShortcut("Ctrl+S")

        self.saveFileAs = QAction(QIcon(os.path.join(self.imagePath, 'assets', 'file-solid.svg')), 'Speichern unter...  (Strg + Shift + S)', self)        
        self.saveFileAs.triggered.connect(self.performSaveFileAs)
        self.saveFileAs.setShortcut("Ctrl+Shift+S")

        self.addBatch = QAction(QIcon(os.path.join(self.imagePath, 'assets', 'plus-solid.svg')), 'Charge hinzufügen  (Strg + A)', self)        
        self.addBatch.triggered.connect(self.performAddBatch)
        self.addBatch.setShortcut("Ctrl+A") 

        self.generateSiloListsButton = QAction(QIcon(os.path.join(self.imagePath, 'assets', 'arrows-rotate-solid.svg')), 'Silo-Listen erstellen  (Strg + G)', self)        
        self.generateSiloListsButton.triggered.connect(self.generateSiloLists)
        self.generateSiloListsButton.setShortcut("Ctrl+G")   

        self.printPlans = QAction(QIcon(os.path.join(self.imagePath, 'assets', 'print-solid.svg')), 'Plan drucken  (Strg + P)', self)        
        #self.printPlans.triggered.connect(self.performPrintPlans)
        self.printPlans.setShortcut("Ctrl+P")     

        self.changeSettings = QAction(QIcon(os.path.join(self.imagePath, 'assets', 'gear-solid.svg')), 'Einstellungen (Strg + E)', self)        
        self.changeSettings.triggered.connect(self.displaySettings)
        self.changeSettings.setShortcut("Ctrl+E") 

        self.menubar.addAction(self.openFileDialog)
        self.menubar.addAction(self.saveFile)
        self.menubar.addAction(self.saveFileAs)
        self.menubar.addSeparator()
        self.menubar.addAction(self.addBatch)  
        self.menubar.addAction(self.generateSiloListsButton)
        self.menubar.addSeparator()
        self.menubar.addAction(self.printPlans) 
        self.menubar.addAction(self.changeSettings)

    def _createTabs(self):
        self.tabs = QTabWidget()
        self.tabExtruder1 = QWidget()
        self.tabExtruder2 = QWidget()
        self.tabHomogenisation = QWidget()
        self.tabSilo = QWidget() 

        self.tabs.addTab(self.tabExtruder1, 'Extruder 1')   
        self.tabs.addTab(self.tabExtruder2, 'Extruder 2') 
        self.tabs.addTab(self.tabHomogenisation, 'Homogenisierung') 
        self.tabs.addTab(self.tabSilo, 'Silo')  

        self.tabLayout = QVBoxLayout()
        self.tabLayout.addWidget(self.tabs)    
             
    def _createPlanerViewExtruder1(self):                     
        
        tableHorizontalHeaders = ['KW', 'Schichten', 'Beginn', 'Ende', 'Artikel-Nr.', 'Chargen-Nr.', 'Dispo.-Nr.', 'Kunde', 'Zusatz', 'Labor', 'Abholung', 't', 'Vorlauf']

        self.tableBatchesExtruder1 = QTableWidget() 
        self.tableBatchesExtruder1.verticalHeader().setVisible(False)
        self.tableBatchesExtruder1.setFixedWidth(1067) 
        self.tableBatchesExtruder1.setFixedHeight(500)  
        self.tableBatchesExtruder1.setColumnCount(13)  
        self.tableBatchesExtruder1.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)   
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(0, 38)     
        self.tableBatchesExtruder1.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(2, 80)  
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(3, 80)  
        self.tableBatchesExtruder1.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(5, 100) 
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(6, 100)
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(7, 150)     
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(8, 120)             
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(9, 80) 
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(10, 80)
        self.tableBatchesExtruder1.horizontalHeader().resizeSection(11, 38) 
        self.tableBatchesExtruder1.horizontalHeader().setSectionResizeMode(12, QHeaderView.ResizeMode.ResizeToContents)          
        self.tableBatchesExtruder1.setHorizontalHeaderLabels(tableHorizontalHeaders)  

        self.sortExtruder1byDeliveryDateButton = QPushButton('Sortieren')
        self.sortExtruder1byDeliveryDateButton.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrow-down-short-wide-solid.svg')))
        self.sortExtruder1byDeliveryDateButton.setFixedWidth(100)
        self.sortExtruder1byDeliveryDateButton.clicked.connect(lambda: self.sortExtruderbyDeliveryDateButton(1))                   

        self.moveToExtruder2 = QPushButton('zu Extruder 2')
        self.moveToExtruder2.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrows-turn-to-dots-solid')))
        self.moveToExtruder2.setFixedWidth(100)
        self.moveToExtruder2.clicked.connect(lambda: self.moveBatchToExtruder(1))

        self.deleteBatchExtruder1 = QPushButton('Löschen')
        self.deleteBatchExtruder1.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'trash-solid.svg')))
        self.deleteBatchExtruder1.setFixedWidth(100)
        self.deleteBatchExtruder1.clicked.connect(lambda: self.deleteBatchFromListExtruder(1))

        self.moveRowUp1 = QPushButton('nach oben')
        self.moveRowUp1.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'up-long-solid.svg')))
        self.moveRowUp1.setFixedWidth(100)
        self.moveRowUp1.clicked.connect(lambda: self.moveBatchRowUp(1))

        self.moveRowDown1 = QPushButton('nach unten')
        self.moveRowDown1.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'down-long-solid.svg')))
        self.moveRowDown1.setFixedWidth(100)
        self.moveRowDown1.clicked.connect(lambda: self.moveBatchRowDown(1))

        self.createShiftPlan1 = QPushButton('Schichten')
        self.createShiftPlan1.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'user-clock-solid.svg')))
        self.createShiftPlan1.setFixedWidth(100)
        self.createShiftPlan1.clicked.connect(lambda: self.createShiftPlan(1))

        self.enumerateBatches1 = QPushButton('Nummerieren')
        self.enumerateBatches1.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrow-down-1-9-solid.svg')))
        self.enumerateBatches1.setFixedWidth(100)
        self.enumerateBatches1.clicked.connect(lambda: self.enumerateBatches(1))

        tabLayoutExtruder1 = QVBoxLayout()
        buttonsLayoutExtruder1 = QHBoxLayout() 
        
        buttonsLayoutExtruder1.addWidget(self.moveToExtruder2)  
        buttonsLayoutExtruder1.addWidget(self.sortExtruder1byDeliveryDateButton) 
        buttonsLayoutExtruder1.addWidget(self.moveRowUp1) 
        buttonsLayoutExtruder1.addWidget(self.moveRowDown1)
        buttonsLayoutExtruder1.addWidget(self.createShiftPlan1)
        buttonsLayoutExtruder1.addWidget(self.enumerateBatches1)
        buttonsLayoutExtruder1.addWidget(self.deleteBatchExtruder1) 
        buttonsLayoutExtruder1.addStretch()
        
        tabLayoutExtruder1.addLayout(buttonsLayoutExtruder1) 
        tabLayoutExtruder1.addWidget(self.tableBatchesExtruder1)                  
        tabLayoutExtruder1.addStretch()

        self.tabExtruder1.setLayout(tabLayoutExtruder1)  

    def _createPlanerViewExtruder2(self):  

        tableHorizontalHeaders = ['KW', 'Schichten', 'Beginn', 'Ende', 'Artikel-Nr.', 'Chargen-Nr.', 'Dispo.-Nr.', 'Kunde', 'Zusatz', 'Labor', 'Abholung', 't', 'Vorlauf']    
                        
        self.tableBatchesExtruder2 = QTableWidget() 
        self.tableBatchesExtruder2.verticalHeader().setVisible(False) 
        self.tableBatchesExtruder2.setFixedWidth(1067) 
        self.tableBatchesExtruder2.setFixedHeight(500)    
        self.tableBatchesExtruder2.setColumnCount(13)  
        self.tableBatchesExtruder2.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)     
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(0, 38)     
        self.tableBatchesExtruder2.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(2, 80)  
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(3, 80)  
        self.tableBatchesExtruder2.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(5, 100) 
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(6, 100)
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(7, 150)     
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(8, 120)             
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(9, 80) 
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(10, 80)
        self.tableBatchesExtruder2.horizontalHeader().resizeSection(11, 38) 
        self.tableBatchesExtruder2.horizontalHeader().setSectionResizeMode(12, QHeaderView.ResizeMode.ResizeToContents)   
        self.tableBatchesExtruder2.setHorizontalHeaderLabels(tableHorizontalHeaders)

        self.sortExtruder2byDeliveryDateButton = QPushButton('Sortieren')
        self.sortExtruder2byDeliveryDateButton.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrow-down-short-wide-solid.svg')))
        self.sortExtruder2byDeliveryDateButton.setFixedWidth(100)
        self.sortExtruder2byDeliveryDateButton.clicked.connect(lambda: self.sortExtruderbyDeliveryDateButton(2))                   

        self.moveToExtruder1 = QPushButton('zu Extruder 1')
        self.moveToExtruder1.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrows-turn-to-dots-solid')))
        self.moveToExtruder1.setFixedWidth(100)
        self.moveToExtruder1.clicked.connect(lambda: self.moveBatchToExtruder(2))

        self.deleteBatchExtruder2 = QPushButton('Löschen')
        self.deleteBatchExtruder2.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'trash-solid.svg')))
        self.deleteBatchExtruder2.setFixedWidth(100)
        self.deleteBatchExtruder2.clicked.connect(lambda: self.deleteBatchFromListExtruder(2))

        self.moveRowUp2 = QPushButton('nach oben')
        self.moveRowUp2.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'up-long-solid.svg')))
        self.moveRowUp2.setFixedWidth(100)
        self.moveRowUp2.clicked.connect(lambda: self.moveBatchRowUp(2))

        self.moveRowDown2 = QPushButton('nach unten')
        self.moveRowDown2.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'down-long-solid.svg')))
        self.moveRowDown2.setFixedWidth(100)
        self.moveRowDown2.clicked.connect(lambda: self.moveBatchRowDown(2))

        self.createShiftPlan2 = QPushButton('Schichten')
        self.createShiftPlan2.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'user-clock-solid.svg')))
        self.createShiftPlan2.setFixedWidth(100)
        self.createShiftPlan2.clicked.connect(lambda: self.createShiftPlan(2))

        self.enumerateBatches2 = QPushButton('Nummerieren')
        self.enumerateBatches2.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrow-down-1-9-solid.svg')))
        self.enumerateBatches2.setFixedWidth(100)
        self.enumerateBatches2.clicked.connect(lambda: self.enumerateBatches(2))

        tabLayoutExtruder2 = QVBoxLayout()
        buttonsLayoutExtruder2 = QHBoxLayout() 
        
        buttonsLayoutExtruder2.addWidget(self.moveToExtruder1)
        buttonsLayoutExtruder2.addWidget(self.sortExtruder2byDeliveryDateButton)  
        buttonsLayoutExtruder2.addWidget(self.moveRowUp2) 
        buttonsLayoutExtruder2.addWidget(self.moveRowDown2)         
        buttonsLayoutExtruder2.addWidget(self.createShiftPlan2)
        buttonsLayoutExtruder2.addWidget(self.enumerateBatches2)
        buttonsLayoutExtruder2.addWidget(self.deleteBatchExtruder2) 
        buttonsLayoutExtruder2.addStretch()
        
        tabLayoutExtruder2.addLayout(buttonsLayoutExtruder2) 
        tabLayoutExtruder2.addWidget(self.tableBatchesExtruder2)                  
        tabLayoutExtruder2.addStretch()

        self.tabExtruder2.setLayout(tabLayoutExtruder2)        
    
    def _createPlanerViewHomogenisation(self):                     
        
        tableHorizontalHeaders = ['KW', 'Schichten', 'Beginn', 'Ende', 'Artikel-Nr.', 'Chargen-Nr.', 'Dispo.-Nr.', 'Kunde', 'Zusatz', 'Labor', 'Abholung', 't', 'Vorlauf']

        self.tableBatchesHomogenisation = QTableWidget() 
        self.tableBatchesHomogenisation.verticalHeader().setVisible(False)
        self.tableBatchesHomogenisation.setFixedWidth(1067) 
        self.tableBatchesHomogenisation.setFixedHeight(500)  
        self.tableBatchesHomogenisation.setColumnCount(13)  
        self.tableBatchesHomogenisation.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)   
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(0, 38)     
        self.tableBatchesHomogenisation.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(2, 80)  
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(3, 80)  
        self.tableBatchesHomogenisation.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(5, 100) 
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(6, 100)
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(7, 150)     
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(8, 120)             
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(9, 80) 
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(10, 80)
        self.tableBatchesHomogenisation.horizontalHeader().resizeSection(11, 38) 
        self.tableBatchesHomogenisation.horizontalHeader().setSectionResizeMode(12, QHeaderView.ResizeMode.ResizeToContents)          
        self.tableBatchesHomogenisation.setHorizontalHeaderLabels(tableHorizontalHeaders)  

        self.sortHomogenisationbyDeliveryDateButton = QPushButton('Sortieren')
        self.sortHomogenisationbyDeliveryDateButton.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrow-down-short-wide-solid.svg')))
        self.sortHomogenisationbyDeliveryDateButton.setFixedWidth(100)
        self.sortHomogenisationbyDeliveryDateButton.clicked.connect(lambda: self.sortExtruderbyDeliveryDateButton(3))                

        self.deleteBatchHomogenisation = QPushButton('Löschen')
        self.deleteBatchHomogenisation.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'trash-solid.svg')))
        self.deleteBatchHomogenisation.setFixedWidth(100)
        self.deleteBatchHomogenisation.clicked.connect(lambda: self.deleteBatchFromListExtruder(3))

        self.moveRowUp3 = QPushButton('nach oben')
        self.moveRowUp3.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'up-long-solid.svg')))
        self.moveRowUp3.setFixedWidth(100)
        self.moveRowUp3.clicked.connect(lambda: self.moveBatchRowUp(3))

        self.moveRowDown3 = QPushButton('nach unten')
        self.moveRowDown3.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'down-long-solid.svg')))
        self.moveRowDown3.setFixedWidth(100)
        self.moveRowDown3.clicked.connect(lambda: self.moveBatchRowDown(3))

        self.createShiftPlan3 = QPushButton('Schichten')
        self.createShiftPlan3.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'user-clock-solid.svg')))
        self.createShiftPlan3.setFixedWidth(100)
        self.createShiftPlan3.clicked.connect(lambda: self.createShiftPlan(3))

        tabLayoutHomogenisation = QVBoxLayout()
        buttonsLayoutHomogenisation = QHBoxLayout() 

        buttonsLayoutHomogenisation.addWidget(self.sortHomogenisationbyDeliveryDateButton) 
        buttonsLayoutHomogenisation.addWidget(self.moveRowUp3) 
        buttonsLayoutHomogenisation.addWidget(self.moveRowDown3)
        buttonsLayoutHomogenisation.addWidget(self.createShiftPlan3)
        buttonsLayoutHomogenisation.addWidget(self.deleteBatchHomogenisation) 
        buttonsLayoutHomogenisation.addStretch()
        
        tabLayoutHomogenisation.addLayout(buttonsLayoutHomogenisation) 
        tabLayoutHomogenisation.addWidget(self.tableBatchesHomogenisation)                  
        tabLayoutHomogenisation.addStretch()

        self.tabHomogenisation.setLayout(tabLayoutHomogenisation)

    def _createPlanerViewSilo(self):                     
        
        tableHorizontalHeaders = ['KW', 'Schichten', 'Beginn', 'Ende', 'Artikel-Nr.', 'Chargen-Nr.', 'Dispo.-Nr.', 'Kunde', 'Zusatz', 'Labor', 'Abholung', 't', 'Vorlauf']

        self.tableBatchesSilo = QTableWidget() 
        self.tableBatchesSilo.verticalHeader().setVisible(False)
        self.tableBatchesSilo.setFixedWidth(1067) 
        self.tableBatchesSilo.setFixedHeight(500)  
        self.tableBatchesSilo.setColumnCount(13)  
        self.tableBatchesSilo.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)   
        self.tableBatchesSilo.horizontalHeader().resizeSection(0, 38)     
        self.tableBatchesSilo.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesSilo.horizontalHeader().resizeSection(2, 80)  
        self.tableBatchesSilo.horizontalHeader().resizeSection(3, 80)  
        self.tableBatchesSilo.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.tableBatchesSilo.horizontalHeader().resizeSection(5, 100) 
        self.tableBatchesSilo.horizontalHeader().resizeSection(6, 100)
        self.tableBatchesSilo.horizontalHeader().resizeSection(7, 150)     
        self.tableBatchesSilo.horizontalHeader().resizeSection(8, 120)             
        self.tableBatchesSilo.horizontalHeader().resizeSection(9, 80) 
        self.tableBatchesSilo.horizontalHeader().resizeSection(10, 80)
        self.tableBatchesSilo.horizontalHeader().resizeSection(11, 38) 
        self.tableBatchesSilo.horizontalHeader().setSectionResizeMode(12, QHeaderView.ResizeMode.ResizeToContents)          
        self.tableBatchesSilo.setHorizontalHeaderLabels(tableHorizontalHeaders)  

        self.sortSilobyDeliveryDateButton = QPushButton('Sortieren')
        self.sortSilobyDeliveryDateButton.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'arrow-down-short-wide-solid.svg')))
        self.sortSilobyDeliveryDateButton.setFixedWidth(100)
        self.sortSilobyDeliveryDateButton.clicked.connect(lambda: self.sortExtruderbyDeliveryDateButton(4))                

        self.deleteBatchSilo = QPushButton('Löschen')
        self.deleteBatchSilo.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'trash-solid.svg')))
        self.deleteBatchSilo.setFixedWidth(100)
        self.deleteBatchSilo.clicked.connect(lambda: self.deleteBatchFromListExtruder(4))

        self.moveRowUp4 = QPushButton('nach oben')
        self.moveRowUp4.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'up-long-solid.svg')))
        self.moveRowUp4.setFixedWidth(100)
        self.moveRowUp4.clicked.connect(lambda: self.moveBatchRowUp(4))

        self.moveRowDown4 = QPushButton('nach unten')
        self.moveRowDown4.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'down-long-solid.svg')))
        self.moveRowDown4.setFixedWidth(100)
        self.moveRowDown4.clicked.connect(lambda: self.moveBatchRowDown(4))

        self.createShiftPlan4 = QPushButton('Schichten')
        self.createShiftPlan4.setIcon(QIcon(os.path.join(self.imagePath, 'assets', 'user-clock-solid.svg')))
        self.createShiftPlan4.setFixedWidth(100)
        self.createShiftPlan4.clicked.connect(lambda: self.createShiftPlan(4))

        tabLayoutSilo = QVBoxLayout()
        buttonsLayoutSilo = QHBoxLayout() 

        buttonsLayoutSilo.addWidget(self.sortSilobyDeliveryDateButton) 
        buttonsLayoutSilo.addWidget(self.moveRowUp4) 
        buttonsLayoutSilo.addWidget(self.moveRowDown4)
        buttonsLayoutSilo.addWidget(self.createShiftPlan4)
        buttonsLayoutSilo.addWidget(self.deleteBatchSilo) 
        buttonsLayoutSilo.addStretch()
        
        tabLayoutSilo.addLayout(buttonsLayoutSilo) 
        tabLayoutSilo.addWidget(self.tableBatchesSilo)                  
        tabLayoutSilo.addStretch()

        self.tabSilo.setLayout(tabLayoutSilo)

    def _createMaster(self):       
        

        masterWidget = QWidget()   
        masterWidget.setLayout(self.tabLayout)       

        self.setCentralWidget(masterWidget)          
           
    def performAddBatch(self):
        
        self.tableBatchesExtruder1.setDisabled(True)
        self.tableBatchesExtruder2.setDisabled(True)        
        self.moveToExtruder2.setDisabled(True)
        self.moveToExtruder1.setDisabled(True) 
        self.sortExtruder1byDeliveryDateButton.setDisabled(True)
        self.deleteBatchExtruder1.setDisabled(True)  
        self.moveRowUp1.setDisabled(True)
        self.sortExtruder2byDeliveryDateButton.setDisabled(True)
        self.deleteBatchExtruder2.setDisabled(True)  
        self.moveRowUp2.setDisabled(True)
        self.tabs.setDisabled(True)
        self.menubar.setDisabled(True)


        self.closeMenu = False

        if self.w is None:            
            self.w = AddBatchWindow(self.attr1, self.attr2, self.attrPack, self.attrLab)         
            self.w.show()
            self.w.finished.connect(self.closeSecondaryWindow)
            self.w.added.connect(self.addBatchesToList)

        else:
            self.w.close()
            self.w = None  

    def displaySettings(self):
        
        self.tableBatchesExtruder1.setDisabled(True)
        self.tableBatchesExtruder2.setDisabled(True)        
        self.moveToExtruder2.setDisabled(True)
        self.moveToExtruder1.setDisabled(True) 
        self.sortExtruder1byDeliveryDateButton.setDisabled(True)
        self.deleteBatchExtruder1.setDisabled(True)  
        self.moveRowUp1.setDisabled(True)
        self.sortExtruder2byDeliveryDateButton.setDisabled(True)
        self.deleteBatchExtruder2.setDisabled(True)  
        self.moveRowUp2.setDisabled(True)
        self.tabs.setDisabled(True)
        self.menubar.setDisabled(True)

        self.closeMenu = False

        if self.w is None:            
            self.w = SettingshWindow(self.sortBy, self.timeNormal, self.timeDensity, self.timeMechanics, self.timeReach)         
            self.w.show()
            self.w.finished.connect(self.closeSecondaryWindow)
            self.w.added.connect(self.writeSettingsToIni)

        else:
            self.w.close()
            self.w = None 

    def closeSecondaryWindow(self):
        self.w = None 
        self.tableBatchesExtruder1.setDisabled(False)
        self.tableBatchesExtruder2.setDisabled(False)        
        self.moveToExtruder2.setDisabled(False)
        self.moveToExtruder1.setDisabled(False)
        self.sortExtruder1byDeliveryDateButton.setDisabled(False)
        self.deleteBatchExtruder1.setDisabled(False) 
        self.moveRowUp1.setDisabled(False)
        self.sortExtruder2byDeliveryDateButton.setDisabled(False)
        self.deleteBatchExtruder2.setDisabled(False)
        self.moveRowUp2.setDisabled(False)
        self.tabs.setDisabled(False)
        self.menubar.setDisabled(False)
        
        self.closeMenu = True        

    def addBatchesToList(self, addBatchArray): 
        self.tableBatchesExtruder1.blockSignals(True)
        self.tableBatchesExtruder2.blockSignals(True)

        deliveryDate = datetime.datetime.strptime(addBatchArray[10], '%d.%m.%Y')
        
        if addBatchArray[9] == 0: 
            productionDate = deliveryDate - datetime.timedelta(days=self.timeNormal)  
            
        elif addBatchArray[9] == 1:            
            productionDate = datetime.datetime.strptime(addBatchArray[10], '%d.%m.%Y') - datetime.timedelta(days=self.timeDensity)
            
        elif addBatchArray[9] == 2:            
            productionDate = datetime.datetime.strptime(addBatchArray[10], '%d.%m.%Y') - datetime.timedelta(days=self.timeMechanics)
                                 
        elif addBatchArray[9] == 3:            
            productionDate = datetime.datetime.strptime(addBatchArray[10], '%d.%m.%Y') - datetime.timedelta(days=self.timeReach)
            
        
        timeToDelivery = (deliveryDate - productionDate).days
       
        year =  addBatchArray[10][-2:]              
        
        calendarWeek = productionDate.strftime('%V')

        rowPosition = self.tableBatchesExtruder1.rowCount()
        self.tableBatchesExtruder1.insertRow(rowPosition)

        rx = QRegularExpression("SP\\d{1,9}")
        rx2 = QRegularExpression("32.\\d{1,4}")
        rx3 = QRegularExpression("1-\\d{1,2}-\\d{1,3}")
        rx4 = QRegularExpression("\\d{1,2}")

        item = 0
        for item in range(len(addBatchArray)):
            match item:
                case 0:                    
                    whichCalendarWeek = QLineEdit()
                    whichCalendarWeek.setText(calendarWeek)
                    whichCalendarWeek.setEnabled(False)
                    whichCalendarWeek.setFixedWidth(38)
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichCalendarWeek)

                case 1:                                    
                    whichShift = QComboBox()
                    whichShift.addItems(self.attrShift)
                    whichShift.setProperty('row', rowPosition)
                    whichShift.currentIndexChanged.connect(lambda: self.shiftChanged(1))
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichShift)

                case 2:
                    buttonProductionDate = QDateEdit()
                    buttonProductionDate.setFixedWidth(80)
                    buttonProductionDate.setDate(productionDate) 
                    buttonProductionDate.setProperty('row', rowPosition)       
                    buttonProductionDate.dateChanged.connect(lambda: self.productionStartDateChangedInTable(1))
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, buttonProductionDate) 

                case 3:
                    buttonProductionDate = QDateEdit()
                    buttonProductionDate.setFixedWidth(80)
                    buttonProductionDate.setDate(productionDate)
                    buttonProductionDate.setProperty('row', rowPosition)       
                    buttonProductionDate.dateChanged.connect(lambda: self.productionEndDateChangedInTable(1))
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, buttonProductionDate)                    

                case 4:                                                    
                    whichArticle = QComboBox()
                    whichArticle.addItem('32.')
                    whichArticle.addItems(self.attr2)                    
                    whichArticle.setCurrentIndex(addBatchArray[4])
                    whichArticle.setEditable(True) 
                    whichArticle.setValidator(QRegularExpressionValidator(rx2, self))                    
                    #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichArticle)  

                case 5:                                         
                    newBatchNo = QLineEdit()
                    newBatchNo.setText('1-'+year+'-')
                    newBatchNo.setValidator(QRegularExpressionValidator(rx3, self))
                    newBatchNo.setFixedWidth(100) 
                    newBatchNo.setMaxLength(8)
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, newBatchNo) 

                case 6:                                         
                    newDispo = QLineEdit()
                    newDispo.setText(addBatchArray[item])
                    newDispo.setValidator(QRegularExpressionValidator(rx, self))
                    newDispo.setFixedWidth(100) 
                    newDispo.setMaxLength(8)
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, newDispo)    

                case 7:                                
                    whichCustomer = QComboBox()
                    whichCustomer.addItem(' ')
                    whichCustomer.addItems(self.attr1)
                    whichCustomer.setCurrentIndex(addBatchArray[7])
                    whichCustomer.setEditable(True)
                    #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichCustomer)   

                case 8:                                    
                    whichPackaging = QComboBox()
                    whichPackaging.addItems(self.attrPack)
                    whichPackaging.setCurrentIndex(addBatchArray[8])
                    #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichPackaging)      

                case 9:                                    
                    whichLab = QComboBox()
                    whichLab.addItems(self.attrLab)
                    whichLab.setCurrentIndex(addBatchArray[9]) 
                    whichLab.setProperty('row', rowPosition)
                    whichLab.setFixedWidth(80)
                    whichLab.currentIndexChanged.connect(lambda: self.labChanged(1))
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichLab)             

                case 10:
                    buttonDeliveryDate = QDateEdit()
                    buttonDeliveryDate.setFixedWidth(80)
                    buttonDeliveryDate.setDate(datetime.datetime.strptime(addBatchArray[item], '%d.%m.%Y'))     
                    buttonDeliveryDate.setProperty('row', rowPosition)              
                    buttonDeliveryDate.dateChanged.connect(lambda: self.deliveryDateChangedInTable(1))             
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, buttonDeliveryDate)  
              
        
                case 11:
                    whichBatchSize = QLineEdit()
                    if addBatchArray[11] == '':
                        batchSize = '24'
                    else:
                        batchSize = addBatchArray[11]
                    whichBatchSize.setText(batchSize)
                    whichBatchSize.setEnabled(True)
                    whichBatchSize.setFixedWidth(38)
                    whichBatchSize.setValidator(QRegularExpressionValidator(rx4, self)) 
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichBatchSize)

                case 12:
                    whichDeliveryDate = QLineEdit()
                    whichDeliveryDate.setText(str(timeToDelivery))
                    whichDeliveryDate.setEnabled(False)
                    whichDeliveryDate.setFixedWidth(38)
                    self.tableBatchesExtruder1.setCellWidget(rowPosition, item, whichDeliveryDate)

                    newTimeToDelivery = int(self.tableBatchesExtruder1.cellWidget(rowPosition, 12).text())

                    if self.tableBatchesExtruder1.cellWidget(rowPosition, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                        self.tableBatchesExtruder1.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')
                    elif self.tableBatchesExtruder1.cellWidget(rowPosition, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                        self.tableBatchesExtruder1.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                    elif self.tableBatchesExtruder1.cellWidget(rowPosition, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                        self.tableBatchesExtruder1.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                    elif newTimeToDelivery < self.timeNormal:            
                        self.tableBatchesExtruder1.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')       
                    else:
                        self.tableBatchesExtruder1.cellWidget(rowPosition, 12).setStyleSheet('background-color: white')

                    if self.tableBatchesExtruder1.cellWidget(rowPosition, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):             
                        self.tableBatchesExtruder1.cellWidget(rowPosition, 3).setStyleSheet('background-color: red')
                    else:
                        self.tableBatchesExtruder1.cellWidget(rowPosition, 3).setStyleSheet('background-color: white')            
            

        self.tableBatchesExtruder1.blockSignals(False)
        self.tableBatchesExtruder2.blockSignals(False)

    def writeSettingsToIni(self, changedSettings): 
        self.tableBatchesExtruder1.blockSignals(True)
        self.tableBatchesExtruder2.blockSignals(True)

        self.sortBy = changedSettings[0] 
        match self.sortBy:
            case 0:
                self.sortByColumn = 2
            case 1:
                self.sortByColumn = 3
            case 2:
                self.sortByColumn = 10

        self.timeNormal = int(changedSettings[1])
        self.timeDensity = int(changedSettings[2])
        self.timeMechanics = int(changedSettings[3])
        self.timeReach = int(changedSettings[4])  

        self.config['SETTINGS']['sortby'] = str(self.sortBy)
        self.config['SETTINGS']['timenormal'] = str(changedSettings[1]) 
        self.config['SETTINGS']['timedensity'] = str(changedSettings[2]) 
        self.config['SETTINGS']['timemechanics'] = str(changedSettings[3]) 
        self.config['SETTINGS']['timereach'] = str(changedSettings[4]) 

        with open(os.path.join(self.imagePath, 'settings.ini'), 'w') as configfile:
                self.config.write(configfile) 
            

        self.tableBatchesExtruder1.blockSignals(False)
        self.tableBatchesExtruder2.blockSignals(False)

    def closeEvent(self, event):
        # do stuff
       
        if self.closeMenu == True:
            event.accept()         

        else:
            event.ignore()

    def moveBatchToExtruder(self, table):
        self.tableBatchesExtruder1.blockSignals(True)
        self.tableBatchesExtruder2.blockSignals(True)       

        if table == 1:            
            whichTable = self.tableBatchesExtruder1
            otherTable = self.tableBatchesExtruder2
        else:           
            whichTable = self.tableBatchesExtruder2
            otherTable = self.tableBatchesExtruder1 

        rx = QRegularExpression("SP\\d{1,9}")
        rx2 = QRegularExpression("32.\\d{1,4}")
        rx3 = QRegularExpression("1-\\d{1,2}-\\d{1,3}")
        rx4 = QRegularExpression("\\d{1,2}")
                    
        if len(whichTable.selectionModel().selectedRows()) != 0:
            rowList = []
            
            for row in whichTable.selectionModel().selectedRows():
                rowList.append(row.row())
                
            rowList.sort()    

            for item in rowList:
                
                rowPosition = otherTable.rowCount()
                otherTable.insertRow(rowPosition)

                rowItem = 0
                for rowItem in range(13):                  

                    match rowItem:
                        case 0:                    
                            whichCalendarWeek = QLineEdit()
                            whichCalendarWeek.setText(whichTable.cellWidget(item, rowItem).text())
                            whichCalendarWeek.setEnabled(0)
                            whichCalendarWeek.setFixedWidth(40)
                            otherTable.setCellWidget(rowPosition, rowItem, whichCalendarWeek)

                        case 1:                                            
                            whichShift = QComboBox()
                            whichShift.addItems(self.attrShift)
                            whichShift.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                            if table == 1:
                                whichShift.currentIndexChanged.connect(lambda: self.shiftChanged(2))
                            else:
                                whichShift.currentIndexChanged.connect(lambda: self.shiftChanged(1))
                            whichShift.setProperty('row', rowPosition)                          
                            otherTable.setCellWidget(rowPosition, rowItem, whichShift)                       

                        case 2:                                  
                            buttonProductionDate = QDateEdit()
                            buttonProductionDate.setFixedWidth(80)
                            buttonProductionDate.setDate(datetime.datetime.strptime(whichTable.cellWidget(item, rowItem).text(), '%d.%m.%Y')) 
                            buttonProductionDate.setProperty('row', rowPosition)
                            if table == 1:       
                                buttonProductionDate.dateChanged.connect(lambda: self.productionStartDateChangedInTable(2)) 
                            else: 
                                buttonProductionDate.dateChanged.connect(lambda: self.productionStartDateChangedInTable(1))                                            
                            otherTable.setCellWidget(rowPosition, rowItem, buttonProductionDate) 

                        case 3:                                  
                            buttonProductionDate = QDateEdit()
                            buttonProductionDate.setFixedWidth(80)
                            buttonProductionDate.setDate(datetime.datetime.strptime(whichTable.cellWidget(item, rowItem).text(), '%d.%m.%Y'))
                            buttonProductionDate.setProperty('row', rowPosition)
                            if table == 1:       
                                buttonProductionDate.dateChanged.connect(lambda: self.productionEndDateChangedInTable(2)) 
                            else: 
                                buttonProductionDate.dateChanged.connect(lambda: self.productionEndDateChangedInTable(1))               
                            otherTable.setCellWidget(rowPosition, rowItem, buttonProductionDate)                            

                        case 4:                                
                            whichArticle = QComboBox()
                            whichArticle.addItem('32.')
                            whichArticle.addItems(self.attr2)
                            whichArticle.setValidator(QRegularExpressionValidator(rx2, self))
                            whichArticle.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                            whichArticle.setEditable(True)
                            #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                            otherTable.setCellWidget(rowPosition, rowItem, whichArticle)
                            
                        case 5:                                         
                            newBatchNo = QLineEdit() 
                            if table == 1:                          
                                newBatchNo.setText('2-'+ whichTable.cellWidget(item, 10).text()[-2:] + '-' )
                            else:
                                newBatchNo.setText('1-'+ whichTable.cellWidget(item, 10).text()[-2:] + '-' )
                            newBatchNo.setValidator(QRegularExpressionValidator(rx3, self))
                            newBatchNo.setFixedWidth(100) 
                            newBatchNo.setMaxLength(8)
                            otherTable.setCellWidget(rowPosition, rowItem, newBatchNo) 

                        case 6:                                         
                            newDispo = QLineEdit()
                            newDispo.setText(whichTable.cellWidget(item, rowItem).text())
                            newDispo.setValidator(QRegularExpressionValidator(rx, self))
                            newDispo.setFixedWidth(100) 
                            newDispo.setMaxLength(8)
                            otherTable.setCellWidget(rowPosition, rowItem, newDispo)      

                        case 7:                                
                            whichCustomer = QComboBox()
                            whichCustomer.addItem(' ')
                            whichCustomer.addItems(self.attr1)                                
                            whichCustomer.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                            whichCustomer.setEditable(True)
                            #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                            otherTable.setCellWidget(rowPosition, rowItem, whichCustomer) 

                        case 8:                
                            whichPackaging = QComboBox()
                            whichPackaging.addItems(self.attrPack)
                            whichPackaging.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                            #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                            otherTable.setCellWidget(rowPosition, rowItem, whichPackaging)      

                        case 9:               
                            whichLab = QComboBox()
                            whichLab.addItems(self.attrLab)
                            whichLab.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                            whichLab.setProperty('row', rowPosition)
                            whichLab.setFixedWidth(80)
                            if table == 1:       
                                whichLab.currentIndexChanged.connect(lambda: self.labChanged(2)) 
                            else: 
                                whichLab.currentIndexChanged.connect(lambda: self.labChanged(1))                            
                            otherTable.setCellWidget(rowPosition, rowItem, whichLab)                           
                                
                        case 10:                                  
                            buttonDeliveryDate = QDateEdit()
                            buttonDeliveryDate.setFixedWidth(80)
                            buttonDeliveryDate.setDate(datetime.datetime.strptime(whichTable.cellWidget(item, rowItem).text(), '%d.%m.%Y')) 
                            if table == 1:       
                                buttonDeliveryDate.dateChanged.connect(lambda: self.deliveryDateChangedInTable(2)) 
                            else: 
                                buttonDeliveryDate.dateChanged.connect(lambda: self.deliveryDateChangedInTable(1))   
                            buttonDeliveryDate.setProperty('row', rowPosition)            
                            otherTable.setCellWidget(rowPosition, rowItem, buttonDeliveryDate)

                        case 11:
                            whichBatchSize = QLineEdit()
                            whichBatchSize.setText(whichTable.cellWidget(item, rowItem).text())
                            whichBatchSize.setEnabled(True)
                            whichBatchSize.setFixedWidth(38)
                            whichBatchSize.setValidator(QRegularExpressionValidator(rx4, self)) 
                            otherTable.setCellWidget(rowPosition, rowItem, whichBatchSize)

                        case 12:
                            whichDeliveryDate = QLineEdit()
                            whichDeliveryDate.setText(whichTable.cellWidget(item, rowItem).text())
                            whichDeliveryDate.setEnabled(False)
                            whichDeliveryDate.setFixedWidth(38)
                            otherTable.setCellWidget(rowPosition, rowItem, whichDeliveryDate) 

                            newTimeToDelivery = int(otherTable.cellWidget(rowPosition, rowItem).text() ) 

                            if otherTable.cellWidget(rowPosition, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                                otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')
                            elif otherTable.cellWidget(rowPosition, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                                otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                            elif otherTable.cellWidget(rowPosition, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                                otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                            elif newTimeToDelivery < self.timeNormal:            
                                otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')       
                            else:
                                otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: white')    

                            if otherTable.cellWidget(rowPosition, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):             
                                otherTable.cellWidget(rowPosition, 3).setStyleSheet('background-color: red')
                            else:
                                otherTable.cellWidget(rowPosition, 3).setStyleSheet('background-color: white')                 
            
            for item in rowList:    
                 whichTable.removeRow(item)

        self.tableBatchesExtruder1.blockSignals(False)
        self.tableBatchesExtruder2.blockSignals(False)

    def deleteBatchFromListExtruder(self, table):
        self.tableBatchesExtruder1.blockSignals(True)
        self.tableBatchesExtruder2.blockSignals(True)
        if self.workingOnShiftPlan == False:

            if table == 1:            
                whichTable = self.tableBatchesExtruder1            
            else:           
                whichTable = self.tableBatchesExtruder2
                

            if len(whichTable.selectionModel().selectedRows()) != 0:            

                rowList = []
                for row in whichTable.selectionModel().selectedRows():
                    rowList.append(row.row())

                rowList.sort(reverse=True)            

                for item in rowList:                
                    whichTable.removeRow(item)

            for row in range(whichTable.rowCount()):
                whichTable.cellWidget(row, 1).setProperty('row', row) 
                whichTable.cellWidget(row, 2).setProperty('row', row)           
                whichTable.cellWidget(row, 3).setProperty('row', row)
                whichTable.cellWidget(row, 9).setProperty('row', row) 
                whichTable.cellWidget(row, 10).setProperty('row', row) 

        self.tableBatchesExtruder1.blockSignals(False)
        self.tableBatchesExtruder2.blockSignals(False)
    
    def productionStartDateChangedInTable(self, table):
        if self.workingOnShiftPlan == False:
            self.tableBatchesExtruder1.blockSignals(True)
            self.tableBatchesExtruder2.blockSignals(True)
            changedDate = self.sender()
            row = changedDate.property('row')                

            if table == 1:            
                whichTable = self.tableBatchesExtruder1            
            else:           
                whichTable = self.tableBatchesExtruder2     


            whichShift = whichTable.cellWidget(row, 1).currentText()
            if whichShift == 'N-W-S' or whichShift == 'S-N' or whichShift == 'N-F':          

                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y') + datetime.timedelta(days=1))  
            else:
                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'))

            newKW = datetime.datetime.strptime(whichTable.cellWidget(row, 3).date().toString('dd.MM.yyyy'), '%d.%m.%Y').strftime('%V')
            whichTable.cellWidget(row, 0).setText(newKW)

            newTimeToDelivery = (datetime.datetime.strptime(whichTable.cellWidget(row, 10).date().toString('dd.MM.yyyy'), '%d.%m.%Y') - datetime.datetime.strptime(whichTable.cellWidget(row, 3).date().toString('dd.MM.yyyy'), '%d.%m.%Y')).days
            whichTable.cellWidget(row, 12).setText(str(newTimeToDelivery))

            if whichTable.cellWidget(row, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')
            elif whichTable.cellWidget(row, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif whichTable.cellWidget(row, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')  
            elif newTimeToDelivery < self.timeNormal:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')  
            else:
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: white')

            if whichTable.cellWidget(row, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):             
                whichTable.cellWidget(row, 3).setStyleSheet('background-color: red')
            else:
                whichTable.cellWidget(row, 3).setStyleSheet('background-color: white')
            
            self.tableBatchesExtruder1.blockSignals(False)
            self.tableBatchesExtruder2.blockSignals(False) 

    def productionEndDateChangedInTable(self, table):
        if self.workingOnShiftPlan == False:
            self.tableBatchesExtruder1.blockSignals(True)
            self.tableBatchesExtruder2.blockSignals(True)
            changedDate = self.sender()
            row = changedDate.property('row')
            newKW = datetime.datetime.strptime(changedDate.date().toString('dd.MM.yyyy'), '%d.%m.%Y').strftime('%V')       

            if table == 1:            
                whichTable = self.tableBatchesExtruder1            
            else:           
                whichTable = self.tableBatchesExtruder2
            
            newTimeToDelivery = (datetime.datetime.strptime(whichTable.cellWidget(row, 10).date().toString('dd.MM.yyyy'), '%d.%m.%Y') - datetime.datetime.strptime(changedDate.date().toString('dd.MM.yyyy'), '%d.%m.%Y')).days

            whichTable.cellWidget(row, 0).setText(newKW)

            whichShift = whichTable.cellWidget(row, 1).currentText()
            if whichShift == 'N-W-S' or whichShift == 'S-N' or whichShift == 'N-F':          

                whichTable.cellWidget(row, 2).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y') - datetime.timedelta(days=1))  
            else:
                whichTable.cellWidget(row, 2).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'))

            whichTable.cellWidget(row, 12).setText(str(newTimeToDelivery))
            
            if whichTable.cellWidget(row, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')
            elif whichTable.cellWidget(row, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif whichTable.cellWidget(row, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif newTimeToDelivery < self.timeNormal:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')       
            else:
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: white')

            if whichTable.cellWidget(row, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):             
                whichTable.cellWidget(row, 3).setStyleSheet('background-color: red')
            else:
                whichTable.cellWidget(row, 3).setStyleSheet('background-color: white')
            
            self.tableBatchesExtruder1.blockSignals(False)
            self.tableBatchesExtruder2.blockSignals(False) 

    def deliveryDateChangedInTable(self, table):
        self.tableBatchesExtruder1.blockSignals(True)
        self.tableBatchesExtruder2.blockSignals(True)   
        if self.workingOnShiftPlan == False:     
            changedDate = self.sender()
            row = changedDate.property('row')                   

            if table == 1:            
                whichTable = self.tableBatchesExtruder1            
            else:           
                whichTable = self.tableBatchesExtruder2
            
            newTimeToDelivery = (datetime.datetime.strptime(changedDate.date().toString('dd.MM.yyyy'), '%d.%m.%Y') - datetime.datetime.strptime(whichTable.cellWidget(row, 3).date().toString('dd.MM.yyyy'), '%d.%m.%Y')).days

            whichTable.cellWidget(row, 12).setText(str(newTimeToDelivery))

            if whichTable.cellWidget(row, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')
            elif whichTable.cellWidget(row, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif whichTable.cellWidget(row, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif newTimeToDelivery < self.timeNormal:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')       
            else:
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: white')
        
        self.tableBatchesExtruder1.blockSignals(False)
        self.tableBatchesExtruder2.blockSignals(False) 

    def shiftChanged(self, table): 
        if self.workingOnShiftPlan == False: 
            self.tableBatchesExtruder1.blockSignals(True)
            self.tableBatchesExtruder2.blockSignals(True) 
            row = self.sender().property('row')             
            whichShift = self.sender().currentText()  

            if table == 1:            
                whichTable = self.tableBatchesExtruder1            
            else:           
                whichTable = self.tableBatchesExtruder2 

            newTimeToDelivery = int(whichTable.cellWidget(row, 12).text())        

            if whichShift == 'N-W-S' or whichShift == 'S-N' or whichShift == 'N-F' or whichShift == 'N' or whichShift == 'W-S-N':         

                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y') + datetime.timedelta(days=1))  
                newTimeToDelivery = newTimeToDelivery - 1
                whichTable.cellWidget(row, 12).setText(str(newTimeToDelivery))

            else:
                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'))            

            if whichTable.cellWidget(row, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')
            elif whichTable.cellWidget(row, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif whichTable.cellWidget(row, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif newTimeToDelivery < self.timeNormal:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')       
            else:
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: white')


            if whichTable.cellWidget(row, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):            
                whichTable.cellWidget(row, 3).setStyleSheet('background-color: red')
            else:
                whichTable.cellWidget(row, 3).setStyleSheet('background-color: white')


            self.tableBatchesExtruder1.blockSignals(False)
            self.tableBatchesExtruder2.blockSignals(False)    

    def labChanged(self, table):            
        self.tableBatchesExtruder1.blockSignals(True)
        self.tableBatchesExtruder2.blockSignals(True)
        if self.workingOnShiftPlan == False:

            whichLab = self.sender().currentIndex()
            row = self.sender().property('row')        

            if table == 1:            
                whichTable = self.tableBatchesExtruder1            
            else:           
                whichTable = self.tableBatchesExtruder2

            if whichLab == 0:     
                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y') - datetime.timedelta(days=self.timeNormal))         
            elif whichLab == 1:            
                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y') - datetime.timedelta(days=self.timeDensity))   
            elif whichLab == 2:            
                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y') - datetime.timedelta(days=self.timeMechanics))                               
            elif whichLab == 3:            
                whichTable.cellWidget(row, 3).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y') - datetime.timedelta(days=self.timeReach))

            newTimeToDelivery = (datetime.datetime.strptime(whichTable.cellWidget(row, 10).date().toString('dd.MM.yyyy'), '%d.%m.%Y') - datetime.datetime.strptime(whichTable.cellWidget(row, 3).date().toString('dd.MM.yyyy'), '%d.%m.%Y')).days
            
            newKW = datetime.datetime.strptime(whichTable.cellWidget(row, 3).date().toString('dd.MM.yyyy'), '%d.%m.%Y').strftime('%V')       
            whichTable.cellWidget(row, 0).setText(newKW)

            whichShift = whichTable.cellWidget(row, 1).currentText()

            if whichShift == 'N-W-S' or whichShift == 'S-N' or whichShift == 'N-F':          

                whichTable.cellWidget(row, 2).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y') - datetime.timedelta(days=1))  
            else:
                whichTable.cellWidget(row, 2).setDate(datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'))

            whichTable.cellWidget(row, 12).setText(str(newTimeToDelivery))

            if whichTable.cellWidget(row, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')
            elif whichTable.cellWidget(row, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif whichTable.cellWidget(row, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
            elif newTimeToDelivery < self.timeNormal:            
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')       
            else:
                whichTable.cellWidget(row, 12).setStyleSheet('background-color: white')
        
        self.tableBatchesExtruder1.blockSignals(False)
        self.tableBatchesExtruder2.blockSignals(False) 

    def sortExtruderbyDeliveryDateButton(self, table):
        self.workingOnShiftPlan = True
        if table == 1:            
            whichTable = self.tableBatchesExtruder1
        else:           
            whichTable = self.tableBatchesExtruder2
        

        saveTableDataHelp = {}
        
        for row in range(whichTable.rowCount()): 
            
            saveTableDataHelp[row] = [whichTable.cellWidget(row, 0).text(), whichTable.cellWidget(row, 1).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'), datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 4).currentIndex(), whichTable.cellWidget(row, 5).text(), whichTable.cellWidget(row, 6).text(), whichTable.cellWidget(row, 7).currentIndex(), whichTable.cellWidget(row, 8).currentIndex(), whichTable.cellWidget(row, 9).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 11).text(), whichTable.cellWidget(row, 12).text()]

        saveTableData = dict(sorted(saveTableDataHelp.items(), key=lambda item: item[1][self.sortByColumn]))    
          
        row = 0
        for key in saveTableData:  
            
            for rowItem in range(13):               
                
                match rowItem:
                    case 0:                  

                        whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])

                    case 1:                                           
                        
                        whichTable.cellWidget(row, rowItem).setProperty('row', row)                          
                        whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])                      

                    case 2:
                        whichTable.cellWidget(row, rowItem).setProperty('row', row) 
                        whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                    case 3:                                  
                         
                        whichTable.cellWidget(row, rowItem).setProperty('row', row)                  
                        whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])  

                        if whichTable.cellWidget(row, rowItem).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):            
                            whichTable.cellWidget(row, rowItem).setStyleSheet('background-color: red')
                        else:
                            whichTable.cellWidget(row, rowItem).setStyleSheet('background-color: white')                      

                    case 4:                                
                        
                        whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
                            
                    case 5:                                         
                       
                        whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])
                        

                    case 6:                                         
                        
                        whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])      

                    case 7:                                
                                                       
                        whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])                        

                    case 8:                
                        
                        whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
   
                    case 9:   
                        whichTable.cellWidget(row, rowItem).setProperty('row', row) 
                        whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
                                                          
                    case 10:                                  
                           
                        whichTable.cellWidget(row, rowItem).setProperty('row', row)            
                        whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                    case 11:
                       
                        whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])

                    case 12:
                        
                        whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])

                        newTimeToDelivery = int(saveTableData[key][rowItem])
                        
                        if whichTable.cellWidget(row, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                            whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')
                        elif whichTable.cellWidget(row, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                            whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
                        elif whichTable.cellWidget(row, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                            whichTable.cellWidget(row, 12).setStyleSheet('background-color: red') 
                        elif newTimeToDelivery < self.timeNormal:            
                            whichTable.cellWidget(row, 12).setStyleSheet('background-color: red')       
                        else:
                            whichTable.cellWidget(row, 12).setStyleSheet('background-color: white')

                        if whichTable.cellWidget(row, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):             
                            whichTable.cellWidget(row, 3).setStyleSheet('background-color: red')
                        else:
                            whichTable.cellWidget(row, 3).setStyleSheet('background-color: white')
                        
                             
            row = row + 1
        self.workingOnShiftPlan = False

    def moveBatchRowUp(self, table):
        self.workingOnShiftPlan = True
        if table == 1:            
            whichTable = self.tableBatchesExtruder1
        else:           
            whichTable = self.tableBatchesExtruder2

        if len(whichTable.selectionModel().selectedRows()) != 0:            

            rowList = []            
            for rowTable in whichTable.selectionModel().selectedRows():
                rowList.append(rowTable.row())

            rowList.sort()
            
            rowCount = whichTable.rowCount()

            firstRowtoOverwrite = rowList[0] - 1

            if firstRowtoOverwrite < 0:
                firstRowtoOverwrite = 0

            saveTableDataSelected = {}
            saveTableDataNotSelected = {}

            for row in range(rowCount): 
                if row in rowList:
                    saveTableDataSelected[row] = [whichTable.cellWidget(row, 0).text(), whichTable.cellWidget(row, 1).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'), datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 4).currentIndex(), whichTable.cellWidget(row, 5).text(), whichTable.cellWidget(row, 6).text(), whichTable.cellWidget(row, 7).currentIndex(), whichTable.cellWidget(row, 8).currentIndex(), whichTable.cellWidget(row, 9).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 11).text(), whichTable.cellWidget(row, 12).text()]
                else:
                    if row > firstRowtoOverwrite or row == firstRowtoOverwrite:
                        saveTableDataNotSelected[row] = [whichTable.cellWidget(row, 0).text(), whichTable.cellWidget(row, 1).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'), datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 4).currentIndex(), whichTable.cellWidget(row, 5).text(), whichTable.cellWidget(row, 6).text(), whichTable.cellWidget(row, 7).currentIndex(), whichTable.cellWidget(row, 8).currentIndex(), whichTable.cellWidget(row, 9).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 11).text(), whichTable.cellWidget(row, 12).text()]
            
            saveTableData = saveTableDataSelected | saveTableDataNotSelected    
            
            whichTable.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
            whichTable.clearSelection()
            row = firstRowtoOverwrite
            selectRows = firstRowtoOverwrite + len(rowList) - 1
            for key in saveTableData:

                if row <= selectRows:
                    whichTable.selectRow(row)  

                for rowItem in range(13):                 
                    
                    match rowItem:
                        case 0:                  

                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])

                        case 1:                                           
                            
                            whichTable.cellWidget(row, rowItem).setProperty('row', row)                          
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])                      

                        case 2:

                            whichTable.cellWidget(row, rowItem).setProperty('row', row) 
                            whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                        case 3:                                  
                            
                            whichTable.cellWidget(row, rowItem).setProperty('row', row)                  
                            whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                        case 4:                                
                            
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
                                
                        case 5:                                         
                        
                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])
                            

                        case 6:                                         
                            
                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])      

                        case 7:                                
                                                        
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])                        

                        case 8:                
                            
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
    
                        case 9:   
                            whichTable.cellWidget(row, rowItem).setProperty('row', row) 
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
                                                            
                        case 10:                                  
                            
                            whichTable.cellWidget(row, rowItem).setProperty('row', row)            
                            whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                        case 11:
                        
                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])

                        case 12:
                            
                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])
                                
                row = row + 1
            whichTable.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        self.workingOnShiftPlan = False

    def moveBatchRowDown(self, table):
        self.workingOnShiftPlan = True
        if table == 1:            
            whichTable = self.tableBatchesExtruder1
        else:           
            whichTable = self.tableBatchesExtruder2

        if len(whichTable.selectionModel().selectedRows()) != 0:            

            rowList = []            
            for rowTable in whichTable.selectionModel().selectedRows():
                rowList.append(rowTable.row())

            rowList.sort()
                
            rowCount = whichTable.rowCount()

            firstRowtoOverwrite = rowList[0]
            if len(rowList) < 2:
                lastRowtoOverwrite = firstRowtoOverwrite + 1
            else:
                lastRowtoOverwrite = rowList[-1] + 1

            if firstRowtoOverwrite < 0:
                firstRowtoOverwrite = 0

            saveTableDataSelected = {}
            saveTableDataNotSelected = {}

            for row in range(rowCount): 
                if row in rowList:
                    saveTableDataSelected[row] = [whichTable.cellWidget(row, 0).text(), whichTable.cellWidget(row, 1).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'), datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 4).currentIndex(), whichTable.cellWidget(row, 5).text(), whichTable.cellWidget(row, 6).text(), whichTable.cellWidget(row, 7).currentIndex(), whichTable.cellWidget(row, 8).currentIndex(), whichTable.cellWidget(row, 9).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 11).text(), whichTable.cellWidget(row, 12).text()]
                else:
                    if row >= firstRowtoOverwrite and row <= lastRowtoOverwrite:
                        saveTableDataNotSelected[row] = [whichTable.cellWidget(row, 0).text(), whichTable.cellWidget(row, 1).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'), datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 4).currentIndex(), whichTable.cellWidget(row, 5).text(), whichTable.cellWidget(row, 6).text(), whichTable.cellWidget(row, 7).currentIndex(), whichTable.cellWidget(row, 8).currentIndex(), whichTable.cellWidget(row, 9).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 11).text(), whichTable.cellWidget(row, 12).text()]
                
            saveTableData = saveTableDataNotSelected | saveTableDataSelected    
                
            whichTable.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
            whichTable.clearSelection()
            row = firstRowtoOverwrite
            selectRows = lastRowtoOverwrite - len(rowList) + 1
            for key in saveTableData:

                if row >= selectRows and row <= lastRowtoOverwrite:
                    whichTable.selectRow(row)  

                for rowItem in range(13):                 
                        
                    match rowItem:
                        case 0:                  

                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])

                        case 1:                                           
                                
                            whichTable.cellWidget(row, rowItem).setProperty('row', row)                          
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])                      

                        case 2:

                            whichTable.cellWidget(row, rowItem).setProperty('row', row) 
                            whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                        case 3:                                  
                                
                            whichTable.cellWidget(row, rowItem).setProperty('row', row)                  
                            whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                        case 4:                                
                                
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
                                    
                        case 5:                                         
                            
                           whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])
                                

                        case 6:                                         
                                
                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])      

                        case 7:                                
                                                            
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])                        

                        case 8:                
                                
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
        
                        case 9:   
                            whichTable.cellWidget(row, rowItem).setProperty('row', row) 
                            whichTable.cellWidget(row, rowItem).setCurrentIndex(saveTableData[key][rowItem])
                                                                
                        case 10:                                  
                                
                            whichTable.cellWidget(row, rowItem).setProperty('row', row)            
                            whichTable.cellWidget(row, rowItem).setDate(saveTableData[key][rowItem])

                        case 11:
                            
                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])

                        case 12:
                                
                            whichTable.cellWidget(row, rowItem).setText(saveTableData[key][rowItem])
                                    
                row = row + 1
            whichTable.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.workingOnShiftPlan = False

    def createShiftPlan(self, table):
        self.workingOnShiftPlan = True

        if table == 1:            
            whichTable = self.tableBatchesExtruder1
        else:           
            whichTable = self.tableBatchesExtruder2

        if whichTable.rowCount() != 0:  

            nextShift = whichTable.cellWidget(0, 1).currentIndex() 
            nextStartDay = whichTable.cellWidget(0, 2).text() 
               
        for row in range(whichTable.rowCount()):
            
            match nextShift:
                case 0:
                    if (row + 1) < whichTable.rowCount():

                        if int(float(whichTable.cellWidget(row+1, 11).text() or 24)) <= 12:
                            whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                            whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                            nextStartDay = whichTable.cellWidget(row+1, 3).text()                         

                            whichTable.cellWidget(row+1, 1).setCurrentIndex(7)
                            nextShift = 7 

                        else:                            
                            whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                            whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                            nextStartDay = whichTable.cellWidget(row+1, 3).text() 

                            if datetime.datetime.strptime(nextStartDay, '%d.%m.%Y').strftime('%A') != 'Thursday':
                                whichTable.cellWidget(row+1, 1).setCurrentIndex(2)                         
                                nextShift = 2 
                            else:
                                whichTable.cellWidget(row+1, 1).setCurrentIndex(3)                         
                                nextShift = 3 


                case 1:
                    if (row + 1) < whichTable.rowCount(): 

                        if int(float(whichTable.cellWidget(row+1, 11).text() or 24)) <= 12:
                            whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                            whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))

                            nextStartDay = whichTable.cellWidget(row+1, 3).text()                         

                            whichTable.cellWidget(row+1, 1).setCurrentIndex(5)
                            nextShift = 5 

                        else:                      

                            whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                            whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))

                            nextStartDay = whichTable.cellWidget(row+1, 3).text()                        

                            if datetime.datetime.strptime(nextStartDay, '%d.%m.%Y').strftime('%A') != 'Thursday':
                                whichTable.cellWidget(row+1, 1).setCurrentIndex(0)                         
                                nextShift = 0 
                            else:
                                whichTable.cellWidget(row+1, 1).setCurrentIndex(4)                                                     
                                nextShift = 4  

                case 2:
                    if (row + 1) < whichTable.rowCount():   

                        if int(float(whichTable.cellWidget(row+1, 11).text() or 24)) <= 12:
                            whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                            whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))

                            nextStartDay = whichTable.cellWidget(row+1, 3).text()                         

                            whichTable.cellWidget(row+1, 1).setCurrentIndex(6)
                            nextShift = 6 

                        else:                     

                            whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                            whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                            nextStartDay = whichTable.cellWidget(row+1, 3).text()                         

                            whichTable.cellWidget(row+1, 1).setCurrentIndex(1)
                            nextShift = 1
                        
                
                case 3:
                    if (row + 1) < whichTable.rowCount():                       

                        whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                        whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                        nextStartDay = whichTable.cellWidget(row+1, 3).text()
                        
                        whichTable.cellWidget(row+1, 1).setCurrentIndex(2)                         
                        nextShift = 2

                case 4:
                    if (row + 1) < whichTable.rowCount():                       

                        whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                        whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                        nextStartDay = whichTable.cellWidget(row+1, 3).text()
                        
                        whichTable.cellWidget(row+1, 1).setCurrentIndex(0)                         
                        nextShift = 0 

                case 5:
                    if (row + 1) < whichTable.rowCount():                       

                        whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                        whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                        nextStartDay = whichTable.cellWidget(row+1, 3).text()                         

                        whichTable.cellWidget(row+1, 1).setCurrentIndex(1)
                        nextShift = 1

                case 6:
                    if (row + 1) < whichTable.rowCount():                       

                        whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                        whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                        nextStartDay = whichTable.cellWidget(row+1, 3).text()                         

                        whichTable.cellWidget(row+1, 1).setCurrentIndex(2)
                        nextShift = 2

                case 7:
                    if (row + 1) < whichTable.rowCount():                       

                        whichTable.cellWidget(row+1, 2).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y'))
                        whichTable.cellWidget(row+1, 3).setDate(datetime.datetime.strptime(nextStartDay, '%d.%m.%Y') + datetime.timedelta(days=1))

                        nextStartDay = whichTable.cellWidget(row+1, 3).text()                         

                        whichTable.cellWidget(row+1, 1).setCurrentIndex(0)
                        nextShift = 0
            
            if (row + 1) < whichTable.rowCount(): 
                newKW = datetime.datetime.strptime(whichTable.cellWidget(row+1, 3).date().toString('dd.MM.yyyy'), '%d.%m.%Y').strftime('%V')  
                whichTable.cellWidget(row+1, 0).setText(newKW)
                newTimeToDelivery = (datetime.datetime.strptime(whichTable.cellWidget(row+1, 10).date().toString('dd.MM.yyyy'), '%d.%m.%Y') - datetime.datetime.strptime(whichTable.cellWidget(row+1, 3).date().toString('dd.MM.yyyy'), '%d.%m.%Y')).days
                whichTable.cellWidget(row+1, 12).setText(str(newTimeToDelivery))
                
                        
                if whichTable.cellWidget(row+1, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                    whichTable.cellWidget(row+1, 12).setStyleSheet('background-color: red')
                elif whichTable.cellWidget(row+1, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                    whichTable.cellWidget(row+1, 12).setStyleSheet('background-color: red') 
                elif whichTable.cellWidget(row+1, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                    whichTable.cellWidget(row+1, 12).setStyleSheet('background-color: red') 
                elif newTimeToDelivery < self.timeNormal:            
                    whichTable.cellWidget(row+1, 12).setStyleSheet('background-color: red')       
                else:
                    whichTable.cellWidget(row+1, 12).setStyleSheet('background-color: white') 

                if whichTable.cellWidget(row+1, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):            
                    whichTable.cellWidget(row+1, 3).setStyleSheet('background-color: red')
                else:
                    whichTable.cellWidget(row+1, 3).setStyleSheet('background-color: white')

                
            
        self.workingOnShiftPlan = False               
    
    def enumerateBatches(self, table):
        if table == 1:            
            whichTable = self.tableBatchesExtruder1
        else:           
            whichTable = self.tableBatchesExtruder2

        if whichTable.rowCount() != 0:            

                rowList = []
                for row in range(whichTable.rowCount()): 

                    stringLength = len(whichTable.cellWidget(row, 5).text())

                    if 0 <= stringLength <= 5:
                        rowList.append(0)
                    elif 5 < stringLength <= 6:
                        rowList.append(int(whichTable.cellWidget(row, 5).text()[-1:]))                  
                    elif 6 < stringLength <= 7:
                        rowList.append(int(whichTable.cellWidget(row, 5).text()[-2:]))                  
                    else:                
                        rowList.append(int(whichTable.cellWidget(row, 5).text()[-3:]))                     

                rowList.sort(reverse=True) 
                
                highestBatchNo = rowList[0] + 1               
                                          
                for row in range(whichTable.rowCount()):

                    stringLength = len(whichTable.cellWidget(row, 5).text())

                    if stringLength <= 5 and highestBatchNo < 10:  
                        whichTable.cellWidget(row, 5).setText(whichTable.cellWidget(row, 5).text() + '00' + str(highestBatchNo))                         
                        highestBatchNo = highestBatchNo + 1
                    elif stringLength <= 5 and highestBatchNo < 100:
                        whichTable.cellWidget(row, 5).setText(whichTable.cellWidget(row, 5).text() + '0' + str(highestBatchNo))                        
                        highestBatchNo = highestBatchNo + 1
                    elif stringLength <= 5 and highestBatchNo < 1000:
                        whichTable.cellWidget(row, 5).setText(whichTable.cellWidget(row, 5).text() + str(highestBatchNo))                        
                        highestBatchNo = highestBatchNo + 1                    

    def loadFile(self):
        self.tableBatchesExtruder1.blockSignals(True)
        self.tableBatchesExtruder2.blockSignals(True) 
        
        fileName = QFileDialog.getOpenFileName(self, 'Speichern unter...', self.saveFilePath, 'Excel-Dateien (*.xlsx)' )        
        if fileName != ([], '') and fileName != ('', ''):           
            
            self.saveFilePath = os.path.join(os.path.dirname(__file__), (fileName[0]))            

            self.config['PATH']['LastSaved'] = self.saveFilePath 
            with open(os.path.join(self.imagePath, 'settings.ini'), 'w') as configfile:
                self.config.write(configfile) 

            shutil.copy(self.saveFilePath, os.path.join(os.path.dirname(__file__), os.path.splitext(os.path.basename(fileName[0]))[0]+'.backup'))

            wb = load_workbook(filename=self.saveFilePath)
            sheets = wb.sheetnames

            sheetsNo = 2 ### modify for homogenisation and silo

            for sheet in range(sheetsNo):

                ws = wb[sheets[sheet]]                

                if sheet == 0:            
                    whichTable = self.tableBatchesExtruder1                                        
                else:           
                    whichTable = self.tableBatchesExtruder2                                   

                rx = QRegularExpression("SP\\d{1,9}")
                rx2 = QRegularExpression("32.\\d{1,4}")
                rx3 = QRegularExpression("1-\\d{1,2}-\\d{1,3}")
                rx4 = QRegularExpression("\\d{1,2}")
                rowPosition = 0
                for row in ws.iter_rows(values_only=True):

                    whichTable.insertRow(rowPosition)   

                    for rowItem in range(13): 

                        match rowItem:                            
                            case 0:                    
                                whichCalendarWeek = QLineEdit()
                                whichCalendarWeek.setText(row[rowItem])
                                whichCalendarWeek.setEnabled(0)
                                whichCalendarWeek.setFixedWidth(40)
                                whichTable.setCellWidget(rowPosition, rowItem, whichCalendarWeek)                                

                            case 1:                                            
                                whichShift = QComboBox()
                                whichShift.addItems(self.attrShift)
                                whichShift.setCurrentIndex(row[rowItem])
                                if sheet == 0:
                                    whichShift.currentIndexChanged.connect(lambda: self.shiftChanged(1))
                                else:
                                    whichShift.currentIndexChanged.connect(lambda: self.shiftChanged(2))
                                whichShift.setProperty('row', rowPosition)                          
                                whichTable.setCellWidget(rowPosition, rowItem, whichShift)                       

                            case 2:                                  
                                buttonProductionDate = QDateEdit()
                                buttonProductionDate.setFixedWidth(80)
                                buttonProductionDate.setDate(row[rowItem]) 
                                buttonProductionDate.setProperty('row', rowPosition)
                                if sheet == 0:       
                                    buttonProductionDate.dateChanged.connect(lambda: self.productionStartDateChangedInTable(1)) 
                                else: 
                                    buttonProductionDate.dateChanged.connect(lambda: self.productionStartDateChangedInTable(2))                                            
                                whichTable.setCellWidget(rowPosition, rowItem, buttonProductionDate) 

                            case 3:                                  
                                buttonProductionDate = QDateEdit()
                                buttonProductionDate.setFixedWidth(80)
                                buttonProductionDate.setDate(row[rowItem])
                                buttonProductionDate.setProperty('row', rowPosition)
                                if sheet == 0:       
                                    buttonProductionDate.dateChanged.connect(lambda: self.productionEndDateChangedInTable(1)) 
                                else: 
                                    buttonProductionDate.dateChanged.connect(lambda: self.productionEndDateChangedInTable(2))               
                                whichTable.setCellWidget(rowPosition, rowItem, buttonProductionDate)

                                if row[rowItem].strftime('%Y.%m.%d') <= datetime.datetime.now().strftime('%Y.%m.%d'):           
                                    whichTable.cellWidget(rowPosition, rowItem).setStyleSheet('background-color: red')
                                else:
                                    whichTable.cellWidget(rowPosition, rowItem).setStyleSheet('background-color: white')

                            case 4:                                
                                whichArticle = QComboBox()
                                whichArticle.addItem('32.')
                                whichArticle.addItems(self.attr2)
                                whichArticle.setValidator(QRegularExpressionValidator(rx2, self))
                                whichArticle.setCurrentIndex(row[rowItem])
                                whichArticle.setEditable(True)
                                #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                                whichTable.setCellWidget(rowPosition, rowItem, whichArticle)
                                
                            case 5:                                         
                                newBatchNo = QLineEdit()                                                        
                                newBatchNo.setText(row[rowItem])                                
                                newBatchNo.setValidator(QRegularExpressionValidator(rx3, self))
                                newBatchNo.setFixedWidth(100) 
                                newBatchNo.setMaxLength(8)
                                whichTable.setCellWidget(rowPosition, rowItem, newBatchNo) 

                            case 6:                                         
                                newDispo = QLineEdit()
                                newDispo.setText(row[rowItem])
                                newDispo.setValidator(QRegularExpressionValidator(rx, self))
                                newDispo.setFixedWidth(100) 
                                newDispo.setMaxLength(8)
                                whichTable.setCellWidget(rowPosition, rowItem, newDispo)      

                            case 7:                                
                                whichCustomer = QComboBox()
                                whichCustomer.addItem(' ')
                                whichCustomer.addItems(self.attr1)                                
                                whichCustomer.setCurrentIndex(row[rowItem])
                                #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                                whichCustomer.setEditable(True)
                                whichTable.setCellWidget(rowPosition, rowItem, whichCustomer) 

                            case 8:                
                                whichPackaging = QComboBox()
                                whichPackaging.addItems(self.attrPack)
                                whichPackaging.setCurrentIndex(row[rowItem])
                                #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                                whichTable.setCellWidget(rowPosition, rowItem, whichPackaging)      

                            case 9:               
                                whichLab = QComboBox()
                                whichLab.addItems(self.attrLab)
                                whichLab.setCurrentIndex(row[rowItem])
                                whichLab.setProperty('row', rowPosition)
                                whichLab.setFixedWidth(80)
                                if sheet == 0:       
                                    whichLab.currentIndexChanged.connect(lambda: self.labChanged(1)) 
                                else: 
                                    whichLab.currentIndexChanged.connect(lambda: self.labChanged(2))                         
                                whichTable.setCellWidget(rowPosition, rowItem, whichLab)                           
                                    
                            case 10:                                  
                                buttonDeliveryDate = QDateEdit()
                                buttonDeliveryDate.setFixedWidth(80)
                                buttonDeliveryDate.setDate(row[rowItem]) 
                                if sheet == 0:       
                                    buttonDeliveryDate.dateChanged.connect(lambda: self.deliveryDateChangedInTable(1)) 
                                else: 
                                    buttonDeliveryDate.dateChanged.connect(lambda: self.deliveryDateChangedInTable(2))   
                                buttonDeliveryDate.setProperty('row', rowPosition)            
                                whichTable.setCellWidget(rowPosition, rowItem, buttonDeliveryDate)

                            case 11:
                                whichBatchSize = QLineEdit()
                                whichBatchSize.setText(row[rowItem])
                                whichBatchSize.setEnabled(True)
                                whichBatchSize.setFixedWidth(38)
                                whichBatchSize.setValidator(QRegularExpressionValidator(rx4, self)) 
                                whichTable.setCellWidget(rowPosition, rowItem, whichBatchSize)

                            case 12:
                                whichDeliveryDate = QLineEdit()
                                whichDeliveryDate.setText(row[rowItem])
                                whichDeliveryDate.setEnabled(False)
                                whichDeliveryDate.setFixedWidth(38)
                                whichTable.setCellWidget(rowPosition, rowItem, whichDeliveryDate)  

                                newTimeToDelivery = int(row[rowItem])                               
                        
                                if whichTable.cellWidget(rowPosition, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                                    whichTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')
                                elif whichTable.cellWidget(rowPosition, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                                    whichTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                                elif whichTable.cellWidget(rowPosition, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                                    whichTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                                elif newTimeToDelivery < self.timeNormal:            
                                    whichTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')       
                                else:
                                    whichTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: white')

                                if whichTable.cellWidget(rowPosition, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):             
                                    whichTable.cellWidget(rowPosition, 3).setStyleSheet('background-color: red')
                                else:
                                    whichTable.cellWidget(rowPosition, 3).setStyleSheet('background-color: white')

                                
                    
                    rowPosition = rowPosition + 1 

            self.setLoadedFile = True
        self.tableBatchesExtruder1.blockSignals(False)
        self.tableBatchesExtruder2.blockSignals(False) 
    
    def performSaveFileAs(self): 
        fileName = QFileDialog.getSaveFileName(self, 'Speichern unter...', self.saveFilePath, 'Excel-Dateien (*.xlsx)' )        
        if fileName != ([], '') and fileName != ('', ''):          
            wb = Workbook() 

            ws = wb.active  
            ws.title = 'Extruder 1'
            wb.create_sheet('Extruder 2')
            wb.create_sheet('Homogenisierung')
            wb.create_sheet('Silo')

            sheets = wb.sheetnames

            sheetsNo = 2 ### modify for homogenisation and silo

            for sheet in range(sheetsNo):

                ws = wb[sheets[sheet]]
                if sheet == 0:            
                    whichTable = self.tableBatchesExtruder1                                        
                else:           
                    whichTable = self.tableBatchesExtruder2

                saveTableData = []
                
                saveRow = []        
                for row in range(whichTable.rowCount()): 
                    saveRow = [whichTable.cellWidget(row, 0).text(), whichTable.cellWidget(row, 1).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'), datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 4).currentIndex(), whichTable.cellWidget(row, 5).text(), whichTable.cellWidget(row, 6).text(), whichTable.cellWidget(row, 7).currentIndex(), whichTable.cellWidget(row, 8).currentIndex(), whichTable.cellWidget(row, 9).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 11).text(), whichTable.cellWidget(row, 12).text()]
                    saveTableData.append(saveRow)

                for writeRow in saveTableData:
                    ws.append(writeRow)        

            self.saveFilePath = os.path.join(os.path.dirname(__file__), (fileName[0]))      
            
            wb.save(self.saveFilePath)

            self.config['PATH']['LastSaved'] = self.saveFilePath 
            with open(os.path.join(self.imagePath, 'settings.ini'), 'w') as configfile:
                self.config.write(configfile)
        
    def performSaveFile(self):        

        if self.saveFilePath == '' or self.setLoadedFile == False:
            self.performSaveFileAs()            
        else:
            wb = Workbook() 

            ws = wb.active  
            ws.title = 'Extruder 1'
            wb.create_sheet('Extruder 2')
            wb.create_sheet('Homogenisierung')
            wb.create_sheet('Silo')

            sheets = wb.sheetnames

            sheetsNo = 2 ### modify for homogenisation and silo

            for sheet in range(sheetsNo):

                ws = wb[sheets[sheet]]
                if sheet == 0:            
                    whichTable = self.tableBatchesExtruder1                                        
                else:           
                    whichTable = self.tableBatchesExtruder2

                saveTableData = []
                
                saveRow = []        
                for row in range(whichTable.rowCount()): 
                    saveRow = [whichTable.cellWidget(row, 0).text(), whichTable.cellWidget(row, 1).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 2).text(), '%d.%m.%Y'), datetime.datetime.strptime(whichTable.cellWidget(row, 3).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 4).currentIndex(), whichTable.cellWidget(row, 5).text(), whichTable.cellWidget(row, 6).text(), whichTable.cellWidget(row, 7).currentIndex(), whichTable.cellWidget(row, 8).currentIndex(), whichTable.cellWidget(row, 9).currentIndex(), datetime.datetime.strptime(whichTable.cellWidget(row, 10).text(), '%d.%m.%Y'), whichTable.cellWidget(row, 11).text(), whichTable.cellWidget(row, 12).text()]
                    saveTableData.append(saveRow)

                for writeRow in saveTableData:
                    ws.append(writeRow)
                
                wb.save(self.saveFilePath)    
    
    def generateSiloLists(self):

        extruderList = [self.tableBatchesExtruder1, self.tableBatchesExtruder2 ]

        siloList = [self.tableBatchesSilo, self.tableBatchesHomogenisation, ]      

        rx = QRegularExpression("SP\\d{1,9}")
        rx2 = QRegularExpression("32.\\d{1,4}")
        rx3 = QRegularExpression("1-\\d{1,2}-\\d{1,3}")
        rx4 = QRegularExpression("\\d{1,2}")

        for whichTable in extruderList:

            saveHomogenisationTable = []
            saveSiloTable = []

            if whichTable.rowCount() != 0:
            
                for row in range(whichTable.rowCount()):                     

                    if whichTable.cellWidget(row, 8).currentIndex() == 2:         
                    
                        saveSiloTable.append(row)
                    
                    elif whichTable.cellWidget(row, 8).currentIndex() == 3:            
                    
                        saveHomogenisationTable.append(row)

                saveSiloTable.sort()
                saveHomogenisationTable.sort() 

                for silo in range(len(siloList)):

                    if silo == 0:    
                        whichSiloTable = saveSiloTable 
                    else:
                        whichSiloTable = saveHomogenisationTable  

                    otherTable = siloList[silo]                   

                    for item in whichSiloTable:
                        
                        rowPosition = otherTable.rowCount()
                        otherTable.insertRow(rowPosition)

                        rowItem = 0
                        for rowItem in range(13):                  

                            match rowItem:
                                case 0:                    
                                    whichCalendarWeek = QLineEdit()
                                    whichCalendarWeek.setText(whichTable.cellWidget(item, rowItem).text())
                                    whichCalendarWeek.setEnabled(0)
                                    whichCalendarWeek.setFixedWidth(40)
                                    otherTable.setCellWidget(rowPosition, rowItem, whichCalendarWeek)

                                case 1:                                            
                                    whichShift = QComboBox()
                                    whichShift.addItems(self.attrShift)
                                    whichShift.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                                    if silo == 0:
                                        whichShift.currentIndexChanged.connect(lambda: self.shiftChanged(4))
                                    else:
                                        whichShift.currentIndexChanged.connect(lambda: self.shiftChanged(3))
                                    whichShift.setProperty('row', rowPosition)                          
                                    otherTable.setCellWidget(rowPosition, rowItem, whichShift)                       

                                case 2:                                  
                                    buttonProductionDate = QDateEdit()
                                    buttonProductionDate.setFixedWidth(80)
                                    buttonProductionDate.setDate(datetime.datetime.strptime(whichTable.cellWidget(item, rowItem).text(), '%d.%m.%Y')) 
                                    buttonProductionDate.setProperty('row', rowPosition)
                                    if silo == 0:       
                                        buttonProductionDate.dateChanged.connect(lambda: self.productionStartDateChangedInTable(4)) 
                                    else: 
                                        buttonProductionDate.dateChanged.connect(lambda: self.productionStartDateChangedInTable(3))                                            
                                    otherTable.setCellWidget(rowPosition, rowItem, buttonProductionDate) 

                                case 3:                                  
                                    buttonProductionDate = QDateEdit()
                                    buttonProductionDate.setFixedWidth(80)
                                    buttonProductionDate.setDate(datetime.datetime.strptime(whichTable.cellWidget(item, rowItem).text(), '%d.%m.%Y'))
                                    buttonProductionDate.setProperty('row', rowPosition)
                                    if silo == 0:       
                                        buttonProductionDate.dateChanged.connect(lambda: self.productionEndDateChangedInTable(4)) 
                                    else: 
                                        buttonProductionDate.dateChanged.connect(lambda: self.productionEndDateChangedInTable(3))               
                                    otherTable.setCellWidget(rowPosition, rowItem, buttonProductionDate)                            

                                case 4:                                
                                    whichArticle = QComboBox()
                                    whichArticle.addItem('32.')
                                    whichArticle.addItems(self.attr2)
                                    whichArticle.setValidator(QRegularExpressionValidator(rx2, self))
                                    whichArticle.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                                    whichArticle.setEditable(True)
                                    #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                                    otherTable.setCellWidget(rowPosition, rowItem, whichArticle)
                                    
                                case 5:                                         
                                    newBatchNo = QLineEdit()                                                             
                                    newBatchNo.setText(whichTable.cellWidget(item, rowItem).text())                                    
                                    newBatchNo.setValidator(QRegularExpressionValidator(rx3, self))
                                    newBatchNo.setFixedWidth(100) 
                                    newBatchNo.setMaxLength(8)
                                    otherTable.setCellWidget(rowPosition, rowItem, newBatchNo) 

                                case 6:                                         
                                    newDispo = QLineEdit()
                                    newDispo.setText(whichTable.cellWidget(item, rowItem).text())
                                    newDispo.setValidator(QRegularExpressionValidator(rx, self))
                                    newDispo.setFixedWidth(100) 
                                    newDispo.setMaxLength(8)
                                    otherTable.setCellWidget(rowPosition, rowItem, newDispo)      

                                case 7:                                
                                    whichCustomer = QComboBox()
                                    whichCustomer.addItem(' ')
                                    whichCustomer.addItems(self.attr1)                                
                                    whichCustomer.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                                    whichCustomer.setEditable(True)
                                    #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                                    otherTable.setCellWidget(rowPosition, rowItem, whichCustomer) 

                                case 8:                
                                    whichPackaging = QComboBox()
                                    whichPackaging.addItems(self.attrPack)
                                    whichPackaging.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                                    #whichPackaging.currentIndexChanged.connect(lambda: self.labChanged(1))
                                    otherTable.setCellWidget(rowPosition, rowItem, whichPackaging)      

                                case 9:               
                                    whichLab = QComboBox()
                                    whichLab.addItems(self.attrLab)
                                    whichLab.setCurrentIndex(whichTable.cellWidget(item, rowItem).currentIndex())
                                    whichLab.setProperty('row', rowPosition)
                                    whichLab.setFixedWidth(80)
                                    if silo == 0:       
                                        whichLab.currentIndexChanged.connect(lambda: self.labChanged(4)) 
                                    else: 
                                        whichLab.currentIndexChanged.connect(lambda: self.labChanged(3))                            
                                    otherTable.setCellWidget(rowPosition, rowItem, whichLab)                           
                                        
                                case 10:                                  
                                    buttonDeliveryDate = QDateEdit()
                                    buttonDeliveryDate.setFixedWidth(80)
                                    buttonDeliveryDate.setDate(datetime.datetime.strptime(whichTable.cellWidget(item, rowItem).text(), '%d.%m.%Y')) 
                                    if silo == 0:       
                                        buttonDeliveryDate.dateChanged.connect(lambda: self.deliveryDateChangedInTable(4)) 
                                    else: 
                                        buttonDeliveryDate.dateChanged.connect(lambda: self.deliveryDateChangedInTable(3))   
                                    buttonDeliveryDate.setProperty('row', rowPosition)            
                                    otherTable.setCellWidget(rowPosition, rowItem, buttonDeliveryDate)

                                case 11:
                                    whichBatchSize = QLineEdit()
                                    whichBatchSize.setText(whichTable.cellWidget(item, rowItem).text())
                                    whichBatchSize.setEnabled(True)
                                    whichBatchSize.setFixedWidth(38)
                                    whichBatchSize.setValidator(QRegularExpressionValidator(rx4, self)) 
                                    otherTable.setCellWidget(rowPosition, rowItem, whichBatchSize)

                                case 12:
                                    whichDeliveryDate = QLineEdit()
                                    whichDeliveryDate.setText(whichTable.cellWidget(item, rowItem).text())
                                    whichDeliveryDate.setEnabled(False)
                                    whichDeliveryDate.setFixedWidth(38)
                                    otherTable.setCellWidget(rowPosition, rowItem, whichDeliveryDate) 

                                    newTimeToDelivery = int(otherTable.cellWidget(rowPosition, rowItem).text() ) 

                                    if otherTable.cellWidget(rowPosition, 9).currentIndex() == 1 and newTimeToDelivery < self.timeDensity:            
                                        otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')
                                    elif otherTable.cellWidget(rowPosition, 9).currentIndex() == 2 and newTimeToDelivery < self.timeMechanics:            
                                        otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                                    elif otherTable.cellWidget(rowPosition, 9).currentIndex() == 3 and newTimeToDelivery < self.timeReach:            
                                        otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red') 
                                    elif newTimeToDelivery < self.timeNormal:            
                                        otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: red')       
                                    else:
                                        otherTable.cellWidget(rowPosition, 12).setStyleSheet('background-color: white')    

                                    if otherTable.cellWidget(rowPosition, 3).date().toString('yyyy.MM.dd') <= datetime.datetime.now().strftime('%Y.%m.%d'):             
                                        otherTable.cellWidget(rowPosition, 3).setStyleSheet('background-color: red')
                                    else:
                                        otherTable.cellWidget(rowPosition, 3).setStyleSheet('background-color: white')                    
            
            

def main():               
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), 'assets', 'book-solid.svg')))
    window = MainWindow()
    window.show()
    app.processEvents()
    app.exec()

if __name__ == "__main__":
    main()