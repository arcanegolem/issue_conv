
from PyQt5 import Qt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

from animatedtoggle import AnimatedToggle

import os
import subprocess
import shutil
import sys

class converter(QMainWindow):

    articles = {}
    associations_folder_abspath : str
    timer = QTimer()
    check_timer = QTimer()

    def __init__(self):
        super(converter, self).__init__()

        self.WIDTH = 690
        self.HEIGHT = 540

        self.resize(self.WIDTH, self.HEIGHT)

        # label font
        label_font = QFont("Arial", 11)
        label_font.setBold(True)
        #-------------------------------------------------

        # button font
        button_font = QFont("Arial", 9)
        button_font.setBold(True)
        #-------------------------------------------------

        # main widget
        self.centralwidget = QWidget(self)
        self.centralwidget.resize(self.WIDTH, self.HEIGHT)
        self.setCentralWidget(self.centralwidget)

        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowOpacity(1)
        
        radius = 8
        self.centralwidget.setStyleSheet(
            """
            background: rgb(51, 51, 51);
            border-top-left-radius:{0}px;
            border-bottom-left-radius:{0}px;
            border-top-right-radius:{0}px;
            border-bottom-right-radius:{0}px;
            """.format(radius)
        )
        #-------------------------------------------------

        # shadow graphics
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(15, 15, 15, 230))
        shadow.setOffset(0, 1)
        #-------------------------------------------------

        # top bar
        self.customTopBar = QWidget(self.centralwidget)
        self.customTopBar.setStyleSheet(
            '''
            background: rgb(40, 40, 40);
            border-bottom-left-radius:0px;
            border-bottom-right-radius:0px;
            '''
        )
        self.customTopBar.setGeometry(0, 0, 690, 25)
        self.customTopBar.setGraphicsEffect(shadow)
        #-------------------------------------------------

        # exit button
        exit_button_shadow = QGraphicsDropShadowEffect()
        exit_button_shadow.setBlurRadius(10)
        exit_button_shadow.setColor(QColor(80, 8, 0, 255))
        exit_button_shadow.setOffset(2, 2)

        self.exitButton = QPushButton(self.customTopBar)
        self.exitButton.setStyleSheet(
            """
            QPushButton {background: rgb(190, 25, 0);
                border-top-left-radius:7px;
                border-bottom-left-radius:7px;
                border-top-right-radius:7px;
                border-bottom-right-radius:7px;
            }
            QPushButton:hover {background: rgb(80, 8, 0);
            }
            """
        )
        self.exitButton.setGeometry(670, 5, 15, 15)
        self.exitButton.setGraphicsEffect(exit_button_shadow)
        self.exitButton.clicked.connect(sys.exit)
        #-------------------------------------------------

        # close button
        close_button_shadow = QGraphicsDropShadowEffect()
        close_button_shadow.setBlurRadius(10)
        close_button_shadow.setColor(QColor(9, 22, 79, 255))
        close_button_shadow.setOffset(2, 2)

        self.closeButton = QPushButton(self.customTopBar)
        self.closeButton.setStyleSheet(
            """
            QPushButton {background: rgb(38, 112, 201);
                border-top-left-radius:7px;
                border-bottom-left-radius:7px;
                border-top-right-radius:7px;
                border-bottom-right-radius:7px;
            }
            QPushButton:hover {background: rgb(9, 22, 79);
            }
            """
        )
        self.closeButton.setGeometry(650, 5, 15, 15)
        self.closeButton.setGraphicsEffect(close_button_shadow)
        self.closeButton.clicked.connect(self.showMinimized)
        #-------------------------------------------------

        # file list shadow
        file_list_shadow = QGraphicsDropShadowEffect()
        file_list_shadow.setBlurRadius(15)
        file_list_shadow.setOffset(0, 0)
        file_list_shadow.setColor(QColor(15, 15, 15, 230))
        #-------------------------------------------------

        # file list label
        self.fileLabel = QLabel(self.centralwidget)
        self.fileLabel.setText("Открытые файлы:")
        self.fileLabel.setFont(label_font)
        self.fileLabel.setStyleSheet(
            """
            color: rgb(150, 150, 150);
            padding: 1px;
            """
        )
        self.fileLabel.setGeometry(460, 50, 150, 20)
        #-------------------------------------------------

        # file list
        self.fileList = QListWidget(self.centralwidget)
        self.fileList.setCursor(QCursor(Qt.PointingHandCursor))
        self.fileList.setStyleSheet('''
            QListWidget {background: rgb(46, 46, 46);
                padding: 5px;
                font: 15px;
                }
            QListWidget:item {border-radius: 10px;
                background: rgb(56, 56, 56);
                border-radius: 7px;
                color: rgb(150, 150, 150);
                margin: 2px;
                border-bottom: 2px solid rgb(36, 36, 36);
                border-right: 2px solid rgb(36, 36, 36);
                }
            QListWidget:item:selected {
                background: rgb(40,40,40)
                }
        ''')
        self.fileList.setGeometry(QRect(460, 75, 200, 400))
        self.fileList.setGraphicsEffect(file_list_shadow)
        self.fileList.itemDoubleClicked.connect(self.openTxt)
        #-------------------------------------------------

        # add file button shadow
        add_file_shadow = QGraphicsDropShadowEffect()
        add_file_shadow.setBlurRadius(15)
        add_file_shadow.setOffset(0, 0)
        add_file_shadow.setColor(QColor(15, 15, 15, 230))
        #-------------------------------------------------

        # add file button
        self.addFileButton = QPushButton(self.centralwidget)
        self.addFileButton.setText("Добавить")
        self.addFileButton.setFont(button_font)
        self.addFileButton.setStyleSheet(
            """
            QPushButton {background: rgb(46, 46, 46);
                border-radius: 7px;;
                color: rgb(150, 150, 150);
            }
            QPushButton:hover {background: rgb(56, 56, 56);
                color: rgb(190, 190, 190);
            }
            """
        )
        self.addFileButton.setGeometry(460, 490, 80, 20)
        self.addFileButton.setGraphicsEffect(add_file_shadow)
        self.addFileButton.clicked.connect(self.addFile)
        #-------------------------------------------------

        # delete file button shadow
        delete_file_shadow = QGraphicsDropShadowEffect()
        delete_file_shadow.setBlurRadius(15)
        delete_file_shadow.setOffset(0, 0)
        delete_file_shadow.setColor(QColor(15, 15, 15, 230))
        #-------------------------------------------------

        # delete file button
        self.deleteFileButton = QPushButton(self.centralwidget)
        self.deleteFileButton.setText("Убрать")
        self.deleteFileButton.setFont(button_font)
        self.deleteFileButton.setStyleSheet(
            """
            QPushButton {background: rgb(46, 46, 46);
                border-radius: 7px;;
                color: rgb(150, 150, 150);
            }
            QPushButton:hover {background: rgb(56, 56, 56);
                color: rgb(190, 190, 190);
            }
            """
            )
        self.deleteFileButton.setGeometry(550, 490, 80, 20)
        self.deleteFileButton.setGraphicsEffect(delete_file_shadow)
        self.deleteFileButton.clicked.connect(self.deleteFile)
        #-------------------------------------------------

        # associations list shadow
        associations_list_shadow = QGraphicsDropShadowEffect()
        associations_list_shadow.setBlurRadius(15)
        associations_list_shadow.setOffset(0, 0)
        associations_list_shadow.setColor(QColor(15, 15, 15, 230))
        #-------------------------------------------------

        # associations label
        self.associationsLabel = QLabel(self.centralwidget)
        self.associationsLabel.setText("Доступные таблицы:")
        self.associationsLabel.setFont(label_font)
        self.associationsLabel.setStyleSheet(
            """
            color: rgb(150, 150, 150);
            padding: 1px;
            """
        )
        self.associationsLabel.setGeometry(240, 50, 170, 20)
        #-------------------------------------------------

        # associations list
        self.associationsList = QListWidget(self.centralwidget)
        self.associationsList.setCursor(QCursor(Qt.PointingHandCursor))
        self.associationsList.setStyleSheet('''
            QListWidget {background: rgb(46, 46, 46);
                padding: 5px;
                font: 15px;
                }
            QListWidget:item {border-radius: 10px;
                background: rgb(56, 56, 56);
                border-radius: 7px;
                color: rgb(150, 150, 150);
                margin: 2px;
                border-bottom: 2px solid rgb(36, 36, 36);
                border-right: 2px solid rgb(36, 36, 36);
                }
            QListWidget:item:selected {
                background: rgb(38, 145, 111);
                border-bottom: 2px solid rgb(17, 64, 49);
                border-right: 2px solid rgb(17, 64, 49);
                color: rgb(17, 64, 49);
                }
        ''')
        self.associationsList.setGeometry(QRect(240, 75, 200, 400))
        self.associationsList.setGraphicsEffect(associations_list_shadow)
        #-------------------------------------------------

        # initital parse of associtaions tables folder
        try:
            os.mkdir("Журналы")
        except FileExistsError:
            pass

        self.associations_folder_abspath = os.path.abspath("Журналы")
        
        associations = os.listdir("Журналы")
        associations = filter(lambda x: ".xlsx" in x, associations)

        for item in associations:
            qitem = QListWidgetItem()
            qitem.setText(item)
            qitem.setIcon(QIcon("assets/table.png"))

            self.associationsList.addItem(qitem)
        #-------------------------------------------------

        # add association table button
        add_association_shadow = QGraphicsDropShadowEffect()
        add_association_shadow.setBlurRadius(15)
        add_association_shadow.setOffset(0, 0)
        add_association_shadow.setColor(QColor(15, 15, 15, 230))

        self.addAssociationButton = QPushButton(self.centralwidget)
        self.addAssociationButton.setText("Добавить")
        self.addAssociationButton.setFont(button_font)
        self.addAssociationButton.setStyleSheet(
            """
            QPushButton {background: rgb(46, 46, 46);
                border-radius: 7px;;
                color: rgb(150, 150, 150);
            }
            QPushButton:hover {background: rgb(56, 56, 56);
                color: rgb(190, 190, 190);
            }
            """
            )
        self.addAssociationButton.setGeometry(240, 490, 80, 20)
        self.addAssociationButton.setGraphicsEffect(add_association_shadow)
        self.addAssociationButton.clicked.connect(self.addAssociation)
        #-------------------------------------------------

        # open association folder
        open_association_folder_shadow = QGraphicsDropShadowEffect()
        open_association_folder_shadow.setBlurRadius(15)
        open_association_folder_shadow.setOffset(0, 0)
        open_association_folder_shadow.setColor(QColor(15, 15, 15, 230))

        self.openAssociationsButton = QPushButton(self.centralwidget)
        self.openAssociationsButton.setIcon(QIcon("assets/folder.png"))
        self.openAssociationsButton.setStyleSheet(
            """
            QPushButton {background: rgb(46, 46, 46);
                border-radius: 7px;
                padding: 3
            }
            QPushButton:hover {background: rgb(56, 56, 56);
            }
            """
            )
        self.openAssociationsButton.setGeometry(330, 490, 25, 20)
        self.openAssociationsButton.setGraphicsEffect(open_association_folder_shadow)
        self.openAssociationsButton.clicked.connect(self.openAssociationsFolder)
        #-------------------------------------------------

        # toggle formatting names
        self.toggle = AnimatedToggle(self.centralwidget)
        self.toggle.setGeometry(15, 410, 58, 45)
        #-------------------------------------------------

        # toggle label
        format_label_shadow = QGraphicsDropShadowEffect()
        format_label_shadow.setBlurRadius(10)
        format_label_shadow.setOffset(0, 0)
        format_label_shadow.setColor(QColor(15, 15, 15, 230))

        self.formatLabel = QLabel(self.centralwidget)
        self.formatLabel.setText("Обработка заголовков")
        self.formatLabel.setFont(button_font)
        self.formatLabel.setStyleSheet(
            """
            background: transparent;
            color: rgb(150, 150, 150);
            """
        )
        self.formatLabel.setGeometry(75, 417, 300, 30)
        self.formatLabel.setGraphicsEffect(format_label_shadow)
        #-------------------------------------------------

        # insertion lines
        self.lineTome = QLineEdit(self.centralwidget)
        self.lineTome.setAlignment(Qt.AlignCenter)
        self.lineTome.setStyleSheet(
            """
            QLineEdit {background: rgb(69, 69, 69);
                padding: 3px;
                color: rgb(190, 190, 190);
                border-radius: 12;
                font-weight: bold;
                border: 2px solid rgb(48, 48, 48);
            }
            """
        )
        self.lineTome.setPlaceholderText("Том")
        self.lineTome.setGeometry(20, 325, 200, 25)
        #-------------------------------------------------
        self.lineIssue = QLineEdit(self.centralwidget)
        self.lineIssue.setAlignment(Qt.AlignCenter)
        self.lineIssue.setStyleSheet(
            """
            QLineEdit {background: rgb(69, 69, 69);
                padding: 3px;
                color: rgb(190, 190, 190);
                border-radius: 12;
                font-weight: bold;
                border: 2px solid rgb(48, 48, 48);
            }
            """
        )
        self.lineIssue.setPlaceholderText("Сквозной номер")
        self.lineIssue.setGeometry(20, 355, 200, 25)
        #-------------------------------------------------
        self.linePrice = QLineEdit(self.centralwidget)
        self.linePrice.setAlignment(Qt.AlignCenter)
        self.linePrice.setStyleSheet(
            """
            QLineEdit {background: rgb(69, 69, 69);
                padding: 3px;
                color: rgb(190, 190, 190);
                border-radius: 12;
                font-weight: bold;
                border: 2px solid rgb(48, 48, 48);
            }
            """
        )
        self.linePrice.setPlaceholderText("Цена")
        self.linePrice.setGeometry(20, 385, 200, 25)
        #-------------------------------------------------

        # check timer for convert button
        self.check_timer.timeout.connect(self.checkConvertAvailability)
        self.check_timer.start()
        #-------------------------------------------------

        # status console
        self.statusConsole = QTextBrowser(self.centralwidget)
        self.statusConsole.setFont(QFont("Consolas"))
        self.statusConsole.setStyleSheet(
            """
            background: rgb(204, 204, 204);
            padding: 5px;
            font: 13px;
            font-weight: bold;
            """
        )
        self.statusConsole.setGeometry(20, 75, 200, 215)
        self.statusConsole.append("Test message")
        #-------------------------------------------------

        # convert button
        self.convertButton = QPushButton(self.centralwidget)
        self.convertButton.setText("Конвертировать")
        self.convertButton.setFont(button_font)
        self.convertButton.setIcon(QIcon("assets/convert.png"))
        self.convertButton.setEnabled(False)
        self.convertButton.setStyleSheet(
            """
            QPushButton {background: rgb(0, 102, 0);
                border-radius: 9px;
                color: rgb(0, 204, 102);
            }
            QPushButton:hover {background: rgb(0, 120, 0);
                color: rgb(0, 224, 122);
            }
            QPushButton:disabled {background: rgb(102, 0, 0);
                color: rgb(204, 51, 51)
            }
            """
            )
        self.convertButton.setGeometry(45, 480, 150, 30)
        #-------------------------------------------------

        # success message
        self.messageSuccess = QLabel(self.centralwidget)
        self.messageSuccess.setStyleSheet(
            """
            background: rgb(0, 179, 74);
            border-radius: 7px;
            padding: 8px;
            color: rgb(25, 255, 121);
            font: 12px;
            font-weight: bold;
            """
        )
        self.messageSuccess.setGeometry(20, 40, 10, 10)
        self.messageSuccess.setVisible(False)
        #-------------------------------------------------

        # success message animation
        effect_success = QGraphicsOpacityEffect(self.messageSuccess)
        self.messageSuccess.setGraphicsEffect(effect_success)

        self.messageAnim1 = QPropertyAnimation(self.messageSuccess, b"size")
        self.messageAnim1.setEndValue(QSize(80, 30))
        self.messageAnim1.setDuration(200)
        self.messageAnim2 = QPropertyAnimation(effect_success, b"opacity")
        self.messageAnim2.setStartValue(0)
        self.messageAnim2.setEndValue(1)
        self.messageAnim2.setDuration(400)
        self.messageAnimGroup = QParallelAnimationGroup()
        self.messageAnimGroup.addAnimation(self.messageAnim1)
        self.messageAnimGroup.addAnimation(self.messageAnim2)

        self.messageAnim3 = QPropertyAnimation(effect_success, b"opacity")
        self.messageAnim3.setStartValue(1)
        self.messageAnim3.setEndValue(0)
        self.messageAnim3.setDuration(400)
        #-------------------------------------------------

        # delete message
        self.messageDelete = QLabel(self.centralwidget)
        self.messageDelete.setStyleSheet(
            """
            background: rgb(179, 110, 0);
            border-radius: 7px;
            padding: 8px;
            color: rgb(255, 166, 0);
            font: 12px;
            font-weight: bold;
            """
        )
        self.messageDelete.setGeometry(20, 40, 10, 10)
        self.messageDelete.setVisible(False)
        #-------------------------------------------------

        # delete message animation
        effect_delete = QGraphicsOpacityEffect(self.messageDelete)
        self.messageDelete.setGraphicsEffect(effect_delete)

        self.delMessageAnim1 = QPropertyAnimation(self.messageDelete, b"size")
        self.delMessageAnim1.setEndValue(QSize(80, 30))
        self.delMessageAnim1.setDuration(200)
        self.delMessageAnim2 = QPropertyAnimation(effect_delete, b"opacity")
        self.delMessageAnim2.setStartValue(0)
        self.delMessageAnim2.setEndValue(1)
        self.delMessageAnim2.setDuration(400)
        self.delMessageAnimGroup = QParallelAnimationGroup()
        self.delMessageAnimGroup.addAnimation(self.delMessageAnim1)
        self.delMessageAnimGroup.addAnimation(self.delMessageAnim2)

        self.delMessageAnim3 = QPropertyAnimation(effect_delete, b"opacity")
        self.delMessageAnim3.setStartValue(1)
        self.delMessageAnim3.setEndValue(0)
        self.delMessageAnim3.setDuration(400)
        #-------------------------------------------------


    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.moveFlag = True
            self.movePosition = event.globalPos() - self.pos()
            self.setCursor(QCursor(Qt.OpenHandCursor))
            event.accept()

    def mouseMoveEvent(self, event):
        try:
            if Qt.LeftButton and self.moveFlag:
                self.move(event.globalPos() - self.movePosition)
                self.setCursor(QCursor(Qt.ClosedHandCursor))
                event.accept()
        except AttributeError:
            pass

    def mouseReleaseEvent(self, QMouseEvent):
        self.moveFlag = False
        self.setCursor(Qt.ArrowCursor)


    
    def openTxt(self, item):
        subprocess.Popen(["notepad", self.articles[item.text()]])



    def addFile(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("Text Files (*.txt)")

        file_dialog.exec_()

        files = file_dialog.selectedFiles()

        for file_path in files:
            filename = file_path.split("/")[-1]

            if filename not in self.articles.keys():
                self.articles[filename] = file_path
            
                item = QListWidgetItem()
                item.setText(filename)
                item.setIcon(QIcon("assets/text.png"))

                self.fileList.addItem(item)

                self.successMessage()

    def deleteFile(self):
        selected = self.fileList.selectedItems()

        if not selected:
            return

        for item in selected:
            self.fileList.takeItem(self.fileList.row(item))
            del self.articles[item.text()]

            self.deleteMessage()

    def addAssociation(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("Excel table (*.xlsx)")

        file_dialog.exec_()

        added_associations = file_dialog.selectedFiles()

        for file_path in added_associations:
            destination = self.associations_folder_abspath + "\\" + file_path.split("/")[-1]
            if file_path != self.associations_folder_abspath.replace("\\", "/"):
                shutil.move(file_path, destination)

        self.updateAssociations()

    def updateAssociations(self):
        associations = os.listdir("Журналы")
        associations = filter(lambda x: ".xlsx" in x, associations)

        self.associationsList.clear()

        for item in associations:
            qitem = QListWidgetItem()
            qitem.setText(item)
            qitem.setIcon(QIcon("assets/table.png"))

            self.associationsList.addItem(qitem)

    def openAssociationsFolder(self):
        file_dialog = QFileDialog(directory="Журналы")
        file_dialog.setNameFilter("Excel table (*.xlsx)")

        file_dialog.exec_()

        self.updateAssociations()

    def checkConvertAvailability(self):
        if len(self.lineTome.text()) > 0 and len(self.lineIssue.text()) > 0 and len(self.linePrice.text()) > 0:
            self.convertButton.setEnabled(True)
        else:
            self.convertButton.setEnabled(False)

    def successMessage(self):
        self.messageSuccess.setVisible(True)
        self.messageAnimGroup.start()

        self.timer.singleShot(400, self.successMessageVisText)

        self.timer.singleShot(1000, self.successMessageDisappear)

    def successMessageVisText(self):
        self.messageSuccess.setText("Успешно!")

    def successMessageDisappear(self):
        self.messageAnim3.start()
        
        self.timer.singleShot(800, self.resetSuccessMessage)

    def resetSuccessMessage(self):
        self.messageSuccess.setGeometry(20, 40, 10, 10)
        self.messageSuccess.setText("")



    def deleteMessage(self):
        self.messageDelete.setVisible(True)
        self.delMessageAnimGroup.start()

        self.timer.singleShot(400, self.deleteMessageVisText)

        self.timer.singleShot(1000, self.deleteMessageDisappear)

    def deleteMessageVisText(self):
        self.messageDelete.setText("Удалено!")

    def deleteMessageDisappear(self):
        self.delMessageAnim3.start()

        self.timer.singleShot(800, self.resetDeleteMessage)

    def resetDeleteMessage(self):
        self.messageDelete.setGeometry(20, 40, 10, 10)
        self.messageDelete.setText("")


if __name__ == '__main__':
    app = QApplication([])
    window = converter()
    window.show()
    sys.exit(app.exec_())