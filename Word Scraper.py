import glob
import time

import docx
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout,QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QLCDNumber, QLabel, QWidget, QFileDialog, QListWidget, QListWidgetItem
import sys
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import QSize, QTime, QTimer, Qt

a = chr(34)
file_link = ""
files_links = []
current_folder = False
current_folder_path = ""


class Window(QWidget):
    def __init__(self):
        super().__init__()

        self.setGeometry(272, 72, 800, 600)
        self.setFixedHeight(600)
        self.setFixedWidth(800)
        self.setWindowTitle("\tWord Scraper")
        self.setWindowIcon(QIcon("burger.ico"))


        self.lcd_number()

    def lcd_number(self):

        vbox = QVBoxLayout()
        self.label = QLabel("Word Scraper")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("background-color: #19fc78")
        self.label.setFont(QFont("times new roman", 48))
        self.label.setFixedHeight(120)
        vbox.addWidget(self.label)

        self.label2 = QLabel("      ONLY use OPEN button i.e. file explorer for files, don't use it for selecting folders\n      OR\n"
                             "      Copy the Folder Path and paste in respective box\n"
                             "      Folder Path should be without single/double quotes\n"
                             "      Otherwise the Software won't work")

        self.label2.setStyleSheet("color:red")
        self.label2.setFont(QFont("times new roman", 12))
        self.label2.setFixedHeight(77)

        hbox = QHBoxLayout()

        self.label3 = QLabel(" NOTE  ")
        self.label3.setStyleSheet("color:Red")
        self.label3.setFont(QFont("castellar", 27))
        self.label3.setFixedHeight(72)
        self.label3.setFixedWidth(144)

        hbox.addWidget(self.label3)
        hbox.addWidget(self.label2)

        vbox.addLayout(hbox)



        self.input1 = QLineEdit()
        self.input1.setPlaceholderText("\tEnter the Folder link here")
        self.input1.setFont(QFont("times new roman", 12))
        self.input1.setFixedHeight(60)
        self.input1.setStyleSheet("background-color:white")
        hbox12 = QHBoxLayout()
        hbox12.addWidget(self.input1)

        btn2 = QPushButton(" OPEN ")
        btn2.setFont(QFont("times new roman", 29))
        btn2.setStyleSheet("background-color:yellow")
        btn2.setFixedWidth(120)
        btn2.clicked.connect(self.open)
        hbox12.addWidget(btn2)

        vbox.addLayout(hbox12)

        hbox1 = QHBoxLayout()
        hbox2 = QHBoxLayout()

        self.input2 = QLineEdit()
        self.input2.setPlaceholderText("\tEnter the Keyword here")
        self.input2.setFont(QFont("times new roman", 12))
        self.input2.setFixedHeight(60)
        self.input2.setStyleSheet("background-color:white")
        hbox1.addWidget(self.input2)

        btn1 = QPushButton(" SCAN 1 ")
        btn1.setFont(QFont("times new roman", 36))
        btn1.setStyleSheet("background-color:pink")
        btn1.clicked.connect(self.read_file)
        hbox2.addWidget(btn1)

        btn3 = QPushButton(" SCAN ALL ")
        btn3.setFont(QFont("times new roman", 36))
        btn3.setStyleSheet("background-color:violet")
        btn3.clicked.connect(self.all_files)
        hbox2.addWidget(btn3)

        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)

        self.setLayout(vbox)

    def read_file_controller(self):
        global file_link

        a = chr(34)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        splitted = file_link.split(c)
        file_link = splitted[0] + c + splitted[1] + b + splitted[2]
        self.read_file(self, file_link)

    def open(self):
        global file_link
        path = QFileDialog.getOpenFileName(self, 'Open a file', '',
                                           'All Files (*.*)')
        if path != ('', ''):
            file_link = path[0]
            self.input1.setPlaceholderText(file_link)

    def all_files(self):
        global file_link
        global files_links
        global current_folder
        global current_folder_path
        print(file_link)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        folder_link = self.input1.text().lstrip().rstrip()
        self.keyword = self.input2.text().lstrip().rstrip().lower()

        if current_folder == False:
            if len(file_link) < 1:
                if len(folder_link) < 1:
                    return
                else:
                    folder_link = self.input1.text().lstrip().rstrip()
            else:
                if '.' in file_link:
                    if file_link[2] == b:
                        temp = file_link.split(b)
                        temp2 = temp.pop(len(temp) - 1)
                        folder_link = b.join(temp)
                    elif file_link[2] == c:
                        temp = file_link.split(c)
                        temp2 = temp.pop(len(temp) - 1)
                        folder_link = c.join(temp)

            if '.' in file_link:
                path = folder_link
                self.input1.setPlaceholderText(path)
            elif len(file_link) < 1:
                path = self.input1.text().lstrip().rstrip()
                if path[0] == a:
                    stripper = path.split(a)
                    path = stripper[1]
                elif path[0] == d:
                    stripper = path.split(d)
                    path = stripper[1]
                else:
                    path = self.input1.text().lstrip().rstrip()
            print(path)


            if len(path) < 1:
                return

            path3 = path + '/*.docx'

        else:
            path3 = current_folder_path + '/*.docx'

        files = glob.glob(path3)
        print(files)
        for i in files:
            file_link = i
            files_links.append(i)
        self.read_file()

        message = QMessageBox.question(self, "Choice Message", "Do You want to continue scanning this Folder ? ",
                                       QMessageBox.Yes | QMessageBox.No)

        if message == QMessageBox.Yes:
            if len(current_folder_path) < 1:
                current_folder_path = path
                current_folder = True
            else:
                current_folder_path = current_folder_path
                current_folder = True
        elif message == QMessageBox.No:
            file_link = ""
            folder_link = ""
            current_folder_path = ""
            current_folder = False




    def read_file(self):
        global file_link
        global files_links
        a = chr(34)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        print('1', file_link)
        sub_link = self.input1.text().lstrip().rstrip()
        self.keyword = self.input2.text().lstrip().rstrip().lower()
        if len(file_link) < 1:
            if len(sub_link) < 1:
                return
            else:
                stripper = sub_link.split(a)

                sub_link = stripper[0]

        else:
            sub_link = file_link
            print(sub_link)

        if sub_link.split('.')[-1] != 'docx':
            self.label2.setText("\tFile Type not supported\n\tOnly MS Word (.docx) files are supported\n\tKindly select a MS Word file to Proceed")
            self.label2.setFont(QFont("times new roman", 16))
            return
        print(2, sub_link)
        if len(files_links) < 1:
            files_links.append(sub_link)

        settings_dialog = QDialog()
        settings_dialog.setModal(True)
        settings_dialog.setStyleSheet("background-color:white")
        settings_dialog.setWindowTitle("\ttext file")
        settings_dialog.setGeometry(35, 50, 1300, 660)
        # settings_dialog.showFullScreen()
        vbox_layout = QVBoxLayout()

        self.label = QListWidget()
        num = 0
        text = ""

        for link in files_links:
            self.label.insertItem(num, text)
            self.setFont(QFont("times new roman", 12))
            self.setStyleSheet("background-color:white")

            num += 1

            text = QListWidgetItem()
            name = link.split(b)[-1]
            text.setText("\n" + name + "\n")
            text.setFont(QFont("times new roman", 24))
            self.label.insertItem(num, text)
            num += 1
            doc = docx.Document(link)

            for z, i in enumerate(doc.paragraphs):
                printed = False
                count = i.text.lower().count(self.keyword)
                data = i.text.lower().split()
                word_list = self.keyword.split()
                for word in word_list:
                    address = []
                    for j, index in enumerate(data):
                        if word in index:
                            address.append(j+1)

                    # print(count, "\t#\t algorithm")
                    if len(address) > 0:
                        if printed == False:
                            text = "\n\n\nParagraph " + str(z+1) + "\n" + i.text.lstrip().rstrip()
                            self.label.insertItem(num, text)
                            num += 1
                            printed = True
                        g = ", "
                        address = [str(locat) for locat in address]
                        text = "\n" + "***** " + word + " *****\tfound at these locations\t" + g.join(address)
                        self.label.insertItem(num, text)
                        num += 1
            text = "\n\n\n\n\n"
            self.label.insertItem(num, text)
            num += 1
        vbox_layout.addWidget(self.label)


        settings_dialog.setLayout(vbox_layout)
        settings_dialog.exec_()
        files_links = []



app = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec_())
'''1 E:/copywriting project\kp20220309_revised.docx
E:/copywriting project\kp20220309_revised.docx
2 E:/copywriting project\kp20220309_revised.docx'''
