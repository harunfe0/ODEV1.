import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QTableWidget, QTableWidgetItem, QComboBox, QHBoxLayout
import pandas as pd

class ExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel Veri Yönetimi')
        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()

        self.openButton = QPushButton('Excel Dosyası Aç')
        self.openButton.clicked.connect(self.openFile)
        layout.addWidget(self.openButton)

        self.saveButton = QPushButton('Değişiklikleri Kaydet')
        self.saveButton.clicked.connect(self.saveFile)
        layout.addWidget(self.saveButton)

        self.tableWidget = QTableWidget()
        self.tableWidget.cellChanged.connect(self.onCellChanged)
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)

        self.currentFileName = None

    def openFile(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Aç", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.currentFileName = fileName
            self.loadExcelData(fileName)

    def loadExcelData(self, fileName):
        df = pd.read_excel(fileName)
        if df.empty:
            return

        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                if isinstance(df.iloc[i, j], str):  # Örnek olarak, sadece string tipindeki veriler için dropdown oluşturuluyor
                    comboBox = QComboBox()
                    comboBox.addItem(df.iloc[i, j])  # Mevcut değeri ekleyin
                    comboBox.addItems(["Seçenek 1", "Seçenek 2", "Seçenek 3"])  # Diğer seçenekleri ekleyin
                    comboBox.currentIndexChanged.connect(lambda: self.onComboBoxChanged(i, j, comboBox))
                    self.tableWidget.setCellWidget(i, j, comboBox)
                else:
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))

        self.tableWidget.resizeColumnsToContents()

    def onComboBoxChanged(self, row, column, comboBox):
        # Bu fonksiyon, comboBox'ta bir değişiklik yapıldığında çağrılır
        currentValue = comboBox.currentText()
        print(f"Değer değişti: Satır {row}, Sütun {column}, Yeni Değer: {currentValue}")

    def onCellChanged(self, row, column):
        # Hücredeki değer değiştiğinde çağrılır
        print(f"Hücre değişti: Satır {row}, Sütun {column}")

    def saveFile(self):
        if self.currentFileName:
            df = pd.DataFrame(index=range(self.tableWidget.rowCount()), columns=[self.tableWidget.horizontalHeaderItem(i).text() for i in range(self.tableWidget.columnCount())])

            for i in range(self.tableWidget.rowCount()):
                for j in range(self.tableWidget.columnCount()):
                    if isinstance(self.tableWidget.cellWidget(i, j), QComboBox):
                        df.iloc[i, j] = self.tableWidget.cellWidget(i, j).currentText()
                    else:
                        df.iloc[i, j] = self.tableWidget.item(i, j).text() if self.tableWidget.item(i, j) else ""

            df.to_excel(self.currentFileName, index=False)
            print("Değişiklikler kaydedildi.")

if __name__ == '__main__':
    app = QApplication(sys.argv)  # Burada parantezler eksikti.
    ex = ExcelApp()
    ex.show()
    sys.exit(app.exec_())
