import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QTextEdit, QLineEdit, \
                            QLabel
from PyQt5.QtGui import QIcon, QPixmap

class Extractor(QMainWindow):
    style = '''
                QLineEdit{border-radius: 5px; padding-left: 5px}
                QLabel{font-size: 10px; color: #0996ff}
                QTextEdit{border-radius: 10px; padding: 5px; margin: 0px}
                QMainWindow{background-color: #333333}
                QPushButton{border-radius: 5px; border: solid; border-width: 1px; border-color: #007fff; background-color: none; color: #007fff}
                QPushButton:hover{
                    background-color: #007fff;
                    color: white;
                }
                QPixmap{
                    height: 100px;
                    width: 100px
                }
            '''
    def __init__(self):
        super().__init__()
        self.setGeometry(100, 100, 1150, 650)
        self.setWindowTitle('Data Extractor')
        self.setWindowIcon(QIcon('wde_logo.png'))

        self.InitializeUI()

    def InitializeUI(self):
        self.file_path = QLineEdit(self)
        self.file_path.setToolTip('Enter file path...')
        self.file_path.setGeometry(50, 320, 400, 30)
        self.file_path_req = QLabel(self)
        self.file_path_req.setText('*Excel files only')
        self.file_path_req.move(50, 295)

        self.known_field_label = QLabel(self)
        self.known_field_label.setText('Available data')
        self.known_field_label.move(50, 365)
        self.known_field = QLineEdit(self)
        self.known_field.setGeometry(50, 390, 400, 30)

        self.data_to_extract_label = QLabel(self)
        self.data_to_extract_label.setText('Data to extract')
        self.data_to_extract_label.move(50, 455)
        self.data_to_extract = QLineEdit(self)
        self.data_to_extract.setGeometry(50, 480, 400, 30)

        self.extracted_data = QTextEdit(self)
        self.extracted_data.setGeometry(500, 300, 600, 305)
        self.extracted_data.setPlaceholderText('Search results would display here...')
        self.extracted_data.setStyleSheet('background-color: white')

        self.button = QPushButton(self)
        self.button.setText('Process')
        self.button.setGeometry(50, 570, 100, 35)
        self.button.clicked.connect(self.Process)

        self.logo = QPixmap('de_logo.png')
        self.logo_label = QLabel(self)
        self.logo_label.setGeometry(458, 50, 234, 200)
        self.logo_label.setPixmap(self.logo)

    def Process(self):
        if self.file_path.text() != '':
            file_path = pd.read_excel(f'{self.file_path.text()}')
            #C:\\Users\\PRAISE SUNDAY SABO\\Documents\\Employee Sample Data.xlsx
            data = pd.DataFrame(file_path)
            
            for i in data:
                if self.known_field.text() != None:
                    if self.known_field.text() in list(data[f'{i}']):
                        if self.data_to_extract.text() == '':
                            data.set_index(f'{i}', inplace=True)
                            self.extracted_data.setText(f"{data.loc[f'{self.known_field.text()}']}\n")
                        else:
                            target = list(data[f'{self.data_to_extract.text()}'])[list(data[f'{i}']).index(self.known_field.text())]
                            self.extracted_data.setText(f"You searched for the ' {self.data_to_extract.text()} '  of  ' {self.known_field.text()} ':\n\n{self.data_to_extract.text()}\t{target}")
                            
                else:
                    pass
        else:
            pass

app = QApplication(sys.argv)
app.setStyleSheet(Extractor.style)
window = Extractor()
window.show()
sys.exit(app.exec())