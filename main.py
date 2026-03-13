# Author: Gideon Davis
# Date Created: 03/09/2024

# import statements

import sys
import os
import pytesseract
import subprocess

from PIL import Image
from enum import Enum
import shutil

from openpyxl.workbook import Workbook
from openpyxl.styles import Alignment

from functools import partial

#TODO: FIGURE OUT AN ALTERNATIVE TO PYQT6
#TODO: FUCK RIVERBANK COMPUTING
from PyQt6.QtGui import QPixmap, QTextCursor
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import *

# Check for pytesseract installation
if not (os.path.isfile(r'C:\Program Files\Tesseract-OCR\tesseract.exe')):
    proc = subprocess.Popen([".\\pytesseract\\pytesseract-0.3.10-py3-none-any.whl"], shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

    if (proc.wait() != 0):
        txtOut = "Error: pytesseract failed to install!"

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Function for sorting directory strings
# Sorting algorithm needs improvement
def sort_directory_by_int(dir_list):
    vals = []
    tl = []

    t = 0
    arrlen = len(dir_list)
    for idx in range(arrlen):
        for ch in dir_list[idx]:
            if (ch.isnumeric()) and ch != ' ':
                tl.append(int(ch))
            else:
                break
        sz = len(tl)
        if sz != 0:
            for i in range(0, sz):
                t += tl[sz - i - 1] * pow(10, i)
            vals.append(t)
            t = 0
        tl.clear()
    for i in range(arrlen - 1):
        sort_swap = False
        for j in range(0, arrlen - 1):
            if vals[j] > vals[j + 1]:
                sort_swap = True
                vals[j], vals[j + 1] = vals[j + 1], vals[j]
                dir_list[j], dir_list[j + 1] = dir_list[j + 1], dir_list[j]
        if not sort_swap:
            return


class Window(QMainWindow):

    buttonList = [
                  "Image directory",
                  "Output directory",
                  "Run",
                  "Folder name",
                  "?",
                  "orientation_2_MS",
                  "orientation_1_SM",
                  "Cancel"]
    slider_button = "slider"

    confidenceSelected = 0

    class Orientation(Enum):
        ModelSerial = 0
        SerialModel = 1

    orientation_case = Orientation.ModelSerial

    image_list = [[]]
    image_dir_size = 0

    filename = ''
    new_img_dir = ''
    image_dir = '\0'
    output_dir = '\0'
    output_file_path = '\0'

    def __init__(self):
        super().__init__()
        self.window_min_width, self.window_min_height = 768//2, 768//2
        self.window_max_width, self.window_max_height = 768//2, 768//2
        self.setMinimumSize(self.window_min_width, self.window_min_height)
        self.setMaximumSize(self.window_max_width, self.window_max_height)
        self.setGeometry(0, 0, self.window_max_width, self.window_max_height)
        self.setWindowTitle('Excel Cataloger')
        self.setStyleSheet("font-size: 20px;")

        self.mainWidget = QWidget(self)
        self.setCentralWidget(self.mainWidget)

        self.mainLabel = QLabel(self)
        self.mainWindowBg = QPixmap('UI_BG.png')
        self.mainLabel.setGeometry(0, 0, int(self.window_max_width), int(self.window_max_height))
        self.mainLabel.setScaledContents(True)
        self.mainLabel.setPixmap(self.mainWindowBg)

        self.img_dir_button = QPushButton(self)
        self.textbox_img = QTextEdit(self)

        self.out_dir_button = QPushButton(self)
        self.textbox_out = QTextEdit(self)

        self.output_file_button = QPushButton(self)
        self.textbox_status = QTextEdit(self)

        self.textbox_filename = QTextEdit(self)
        self.textbox_filename_error = QTextEdit(self)

        self.help_button = QPushButton(self)

        self.close_button = QPushButton(self)

        self.orientation_button_SM = QPushButton(self)
        self.orientation_button_MS = QPushButton(self)
        self.textbox_orientation_1 = QTextEdit(self)
        self.textbox_orientation_2 = QTextEdit(self)

        self.confidence_slider = QSlider(self)
        self.textbox_confidence_header = QTextEdit(self)
        self.textbox_confidence_value = QTextEdit(self)

        self.init_ui()

    # Creates all elements in the interface
    def init_ui(self):

        img_dir_y_bound = 36
        img_dir_h_bound = 25
        dir_x_bound = 34
        dir_w_bound = 79
        path_x_bound = 123
        path_w_bound = 159
        out_dir_y_bound = 85
        out_dir_h_bound = img_dir_h_bound
        help_x_bound = 338
        help_y_bound = 10
        help_w_bound = 36
        help_h_bound = 37
        fname_x_bound = 32
        fname_y_bound = 134
        fname_w_bound = 159
        fname_h_bound = 25
        orient_sel_x_bound = 31
        orient_sel_y_bound_SM = 215
        orient_sel_y_bound_MS = 181
        orient_sel_w_bound = 24
        orient_sel_h_bound = 24
        close_x_bound = 28
        close_y_bound = 338
        close_w_bound = 101
        close_h_bound = 24
        run_x_bound = 254
        run_y_bound = close_y_bound
        run_w_bound = close_w_bound
        run_h_bound = close_h_bound
        progress_x_bound = (run_x_bound+close_x_bound)//2
        progress_y_bound = run_y_bound
        progress_w_bound = run_w_bound
        progress_h_bound = run_h_bound
        conf_x_bound = 35
        conf_y_bound = 290
        conf_w_bound = 218
        conf_h_bound = 25

        # Button/Textbox for selecting input directory
        self.img_dir_button.setGeometry(dir_x_bound, img_dir_y_bound, dir_w_bound, img_dir_h_bound)
        self.img_dir_button.setStyleSheet("""
                                          QPushButton {
                                              background-color : rgba(255,255,255,100%);
                                              font-size : 10px;
                                              color: rgba(0,0,0,100%); 
                                              border : 1px solid black;
                                              border-radius: 10px;
                                              alignment : center;
                                          }
                                          QPushButton:hover {
                                              background-color : rgba(217, 230, 250,100%);
                                              border : 1px solid black; 
                                              border-radius: 10px;
                                          }
                                          """)
        self.img_dir_button.setText(self.buttonList[0])
        self.img_dir_button.clicked.connect(partial(self.clicked_btn, 'img_dir_button'))

        # Button/Textbox for selecting output directory
        self.out_dir_button.setGeometry(dir_x_bound, out_dir_y_bound, dir_w_bound, img_dir_h_bound)
        self.out_dir_button.setStyleSheet("""
                                              QPushButton {
                                                  background-color : rgba(255,255,255,100%);
                                                  font-size : 10px;
                                                  color: rgba(0,0,0,100%); 
                                                  border : 1px solid black;
                                                  border-radius: 10px;
                                                  alignment : center;
                                              }
                                              QPushButton:hover {
                                                  background-color : rgba(217, 230, 250,100%);
                                                  border : 1px solid black; 
                                                  border-radius: 10px;
                                              }
                                          """)
        self.out_dir_button.setText(self.buttonList[1])
        self.out_dir_button.clicked.connect(partial(self.clicked_btn, 'output_dir_button'))

        # Button/Textbox for run program
        self.output_file_button.setGeometry(run_x_bound, run_y_bound, run_w_bound, run_h_bound)
        self.output_file_button.setStyleSheet("""
                                                QPushButton {
                                                    background-color : rgba(10,100,200,100%);
                                                    font-size : 12px;
                                                    color: rgba(255,255,255,100%);
                                                    border : 1px solid black;
                                                    border-radius : 10px;
                                                    alignment : center;
                                                }
                                                QPushButton:hover {
                                                    background-color : rgba(10, 100, 255,100%); 
                                                    border : 1px solid white;
                                                    border-radius : 10px; 
                                              }
                                              """)
        self.output_file_button.setText(self.buttonList[2])
        self.output_file_button.clicked.connect(partial(self.clicked_btn, 'output_file_button'))

        # Button/Text for question mark box
        self.help_button.setGeometry(help_x_bound, help_y_bound, help_w_bound, help_h_bound)
        self.help_button.setStyleSheet("""
                                                QPushButton {
                                                    background-color : rgba(100,100,100,100%);
                                                    font-size : 32px;
                                                    color: rgba(0,0,0,100%); 
                                                    border : 1px solid black;
                                                    border-radius : 16px;
                                                    alignment : center;
                                                }
                                                QPushButton:hover {
                                                    background-color : rgba(120, 120, 120,100%);
                                                    border : 1px solid black; 
                                                    border-radius : 16px;
                                              }
                                              """)

        self.help_button.setText(self.buttonList[4])
        self.help_button.clicked.connect(partial(self.clicked_btn, 'help_button'))

        # Texbox for the image directory path
        self.textbox_img.setReadOnly(True)
        self.textbox_img.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_img.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_img.setGeometry(path_x_bound, img_dir_y_bound, path_w_bound, img_dir_h_bound)
        self.textbox_img.setStyleSheet("""
                                          QTextEdit {
                                              background-color : rgba(255,255,255,100%);
                                              font-size : 12px;
                                              color: black;
                                              border : 1px solid black;
                                              border-radius : 10px;
                                              alignment : center;
                                          }
                                      """)
        self.textbox_img.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.textbox_img.setPlaceholderText("Directory Path...")

        # Textbox for output file directory path
        self.textbox_out.setReadOnly(True)
        self.textbox_out.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_out.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_out.setGeometry(path_x_bound, out_dir_y_bound, path_w_bound, out_dir_h_bound)
        self.textbox_out.setStyleSheet("""
                                          QTextEdit { 
                                              background-color : rgba(255,255,255,100%);
                                              font-size : 12px;
                                              color : black;
                                              border : 1px solid black;
                                              border-radius : 9px;
                                              alignment : center;
                                          }
                                       """)
        self.textbox_out.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.textbox_out.setPlaceholderText("Directory Path...")

        # Textbox for current execution status
        self.textbox_status.setReadOnly(True)
        self.textbox_status.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_status.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_status.setGeometry(progress_x_bound, progress_y_bound, progress_w_bound, progress_h_bound)
        self.textbox_status.setStyleSheet("""
                                                 QTextEdit {
                                                         background-color: rgba(255,255,255,0%);
                                                         font-size : 12px;
                                                         color : black;
                                                         border : 0px;
                                                         border-radius : 10px;
                                                         alignment : center;
                                                     }
                                          """)
        self.textbox_status.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)

        # Textbox for output file name
        self.textbox_filename.setReadOnly(False)
        self.textbox_filename.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_filename.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_filename.setGeometry(fname_x_bound, fname_y_bound, fname_w_bound, fname_h_bound)
        self.textbox_filename.setStyleSheet("""
                                                QTextEdit {
                                                    background-color: rgba(255,255,255,100%);
                                                    font-size : 12px;
                                                    color : black;
                                                    border : 1px Solid Black;
                                                    border-radius : 10px;
                                                    alignment : center;
                                                }
                                            """)
        self.textbox_filename.setPlaceholderText("Enter File Name...")

        # Textbox for filename error
        self.textbox_filename_error.setReadOnly(True)
        self.textbox_filename_error.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_filename_error.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_filename_error.setGeometry(fname_x_bound+fname_w_bound+5, fname_y_bound, fname_w_bound, fname_h_bound)
        self.textbox_filename_error.setStyleSheet("""
                                                     QTextEdit {
                                                         background-color: rgba(255,255,255,0%);
                                                         font-size : 12px;
                                                         color : red;
                                                         border : 0px;
                                                         border-radius : 10px;
                                                         alignment : center;
                                                     }
                                                  """)
        self.textbox_filename_error.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)

        # Textbox for orientation selection
        self.textbox_orientation_1.setReadOnly(True)
        self.textbox_orientation_1.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_orientation_1.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_orientation_1.setGeometry(orient_sel_x_bound + 25, orient_sel_y_bound_MS-2, 300, 35)
        self.textbox_orientation_1.setStyleSheet("""
                                                             QTextEdit {
                                                                 background-color: rgba(255,255,255,0%);
                                                                 font-size : 12px;
                                                                 font : italic;
                                                                 color : black;
                                                                 border : 0px;
                                                                 border-radius : 10px;
                                                                 alignment : center;
                                                             }
                                                          """)
        self.textbox_orientation_1.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.textbox_orientation_1.setText("Orientation 1: 'Model.jpg, Serial.jpg'")

        self.textbox_orientation_2.setReadOnly(True)
        self.textbox_orientation_2.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_orientation_2.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_orientation_2.setGeometry(orient_sel_x_bound + 25, orient_sel_y_bound_SM - 2, 300, 35)
        self.textbox_orientation_2.setStyleSheet("""
                                                                     QTextEdit {
                                                                         background-color: rgba(255,255,255,0%);
                                                                         font-size : 12px;
                                                                         font : italic;
                                                                         color : black;
                                                                         border : 0px;
                                                                         border-radius : 10px;
                                                                         alignment : center;
                                                                     }
                                                                  """)
        self.textbox_orientation_2.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.textbox_orientation_2.setText("Orientation 2: 'Serial.jpg, Model.jpg'")

        # orientation bubble select MS
        self.orientation_button_MS.setGeometry(orient_sel_x_bound, orient_sel_y_bound_MS, orient_sel_w_bound, orient_sel_h_bound)
        self.orientation_button_MS.setStyleSheet("""
                                                        QPushButton {
                                                            background-color : rgba(255,255,255,100%);
                                                            font-size : 12px;
                                                            color: rgba(0,0,0,0%); 
                                                            border : 1px solid black;
                                                            border-radius : 12px;
                                                            alignment : center;
                                                        }
                                                        QPushButton:hover {
                                                            background-color : rgba(255, 255, 255,100%);
                                                            border : 3px solid black; 
                                                            border-radius : 12px;
                                                      }
                                                      """)

        self.orientation_button_MS.setText(self.buttonList[6])
        self.orientation_button_MS.clicked.connect(partial(self.clicked_btn, 'orientation_2_MS'))

        # orientation bubble select MS
        self.orientation_button_SM.setGeometry(orient_sel_x_bound, orient_sel_y_bound_SM, orient_sel_w_bound,
                                               orient_sel_h_bound)
        self.orientation_button_SM.setStyleSheet("""
                                                                QPushButton {
                                                                    background-color : rgba(255,255,255,100%);
                                                                    font-size : 12px;
                                                                    color: rgba(0,0,0,0%); 
                                                                    border : 1px solid black;
                                                                    border-radius : 12px;
                                                                    alignment : center;
                                                                }
                                                                QPushButton:hover {
                                                                    background-color : rgba(255, 255, 255,100%);
                                                                    border : 3px solid black; 
                                                                    border-radius : 12px;
                                                              }
                                                              """)

        self.orientation_button_SM.setText(self.buttonList[5])
        self.orientation_button_SM.clicked.connect(partial(self.clicked_btn, 'orientation_1_SM'))

        # Cancel/close button
        self.close_button.setGeometry(close_x_bound, close_y_bound, close_w_bound, close_h_bound)
        self.close_button.setStyleSheet("""
                                                        QPushButton {
                                                            background-color : rgba(10,10,10,100%);
                                                            font-size : 12px;
                                                            color: rgba(255,255,255,100%); 
                                                            border : 1px solid black;
                                                            border-radius : 12px;
                                                            alignment : center;
                                                        }
                                                        QPushButton:hover {
                                                            background-color : rgba(25, 25, 25,100%);
                                                            border : 1px solid white; 
                                                            border-radius : 12px;
                                                      }
                                                      """)

        self.close_button.setText(self.buttonList[7])
        self.close_button.clicked.connect(partial(self.clicked_btn, 'Cancel'))

        # Slider for confidence selection
        self.confidence_slider.setGeometry(conf_x_bound, conf_y_bound, conf_w_bound, conf_h_bound)
        self.confidence_slider.setOrientation(Qt.Orientation.Horizontal)
        self.confidence_slider.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.confidence_slider.setTickPosition(QSlider.TickPosition.TicksBothSides)
        self.confidence_slider.setTickInterval(10)
        self.confidence_slider.setSingleStep(1)
        self.confidence_slider.setMaximum(100)
        self.confidence_slider.setMinimum(0)
        self.confidence_slider.valueChanged.connect(self.slider_event)

        # Textbox for confidence selection slider title

        self.textbox_confidence_header.setReadOnly(True)
        self.textbox_confidence_header.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_confidence_header.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_confidence_header.setGeometry(conf_x_bound, conf_y_bound-conf_h_bound-5, conf_w_bound, conf_h_bound)
        self.textbox_confidence_header.setStyleSheet("""
                                                                     QTextEdit {
                                                                         background-color: rgba(255,255,255,0%);
                                                                         font-size : 12px;
                                                                         font : italic;
                                                                         color : black;
                                                                         border : 0px;
                                                                         border-radius : 10px;
                                                                         alignment : center;
                                                                     }
                                                                  """)
        self.textbox_confidence_header.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.textbox_confidence_header.setText("Select Confidence: 0% - 100%")

        # Textbox for confidence slider value
        self.textbox_confidence_value.setReadOnly(True)
        self.textbox_confidence_value.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_confidence_value.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textbox_confidence_value.setGeometry(conf_x_bound+conf_w_bound+5, conf_y_bound, conf_w_bound//6+3, conf_h_bound)
        self.textbox_confidence_value.setStyleSheet("""
                                                    QTextEdit {
                                                        background-color: rgba(255,255,255,100%);
                                                        font-size : 12px;
                                                        font : bold;
                                                        color : black;
                                                        border : 1px solid black;
                                                        border-radius : 10px;
                                                        alignment : center;
                                                    }
                                                    """)
        self.textbox_confidence_value.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.textbox_confidence_value.setText("0%")

        self.show()

    def clicked_btn(self, value):
        sender = self.sender()

        if sender.text() == self.buttonList[0]:
            self.sel_img_dir()
        elif sender.text() == self.buttonList[1]:
            self.sel_output_dir()
        elif sender.text() == self.buttonList[2]:
            self.process_data()
        elif sender.text() == self.buttonList[4]:
            self.open_manual()
        elif sender.text() == self.buttonList[5]:
            self.orientation_sel(self.Orientation.SerialModel)
        elif sender.text() == self.buttonList[6]:
            self.orientation_sel(self.Orientation.ModelSerial)
        elif sender.text() == self.buttonList[7]:
            sys.exit()
        else:
            print('Invalid Input')

        print(sender.text())

    def slider_event(self):
        self.confidenceSelected = self.confidence_slider.value()
        self.textbox_confidence_value.setText(str(self.confidenceSelected) + "%")

    def sel_img_dir(self):
        print('button 2')
        response = QFileDialog.getExistingDirectory(
            self,
            caption='Select Image Directory'
        )
        self.image_dir = str(response)
        self.textbox_img.moveCursor(QTextCursor.MoveOperation.Start)
        self.textbox_img.setStyleSheet("""
                                                   QTextEdit {
                                                       background-color: rgba(255,255,255,100%);
                                                       font-size : 12px;
                                                       color : black;
                                                       border : 1px Solid Black;
                                                       border-radius : 10px;
                                                       alignment : center;
                                                   }
                                                   """)
        self.textbox_img.clear()
        self.textbox_img.insertPlainText(self.image_dir)
        self.textbox_img.moveCursor(QTextCursor.MoveOperation.End)

    def sel_output_dir(self):
        print('button 1')
        response = QFileDialog.getExistingDirectory(
            self,
            caption='Select Output Directory'
        )
        self.output_dir = str(response)
        self.textbox_out.moveCursor(QTextCursor.MoveOperation.Start)
        self.textbox_out.setStyleSheet("""
                                                    QTextEdit {
                                                        background-color: rgba(255,255,255,100%);
                                                        font-size : 12px;
                                                        color : black;
                                                        border : 1px Solid Black;
                                                        border-radius : 10px;
                                                        alignment : center;
                                                    }
                                                   """)
        self.textbox_out.clear()
        self.textbox_out.insertPlainText(self.output_dir)
        self.textbox_out.moveCursor(QTextCursor.MoveOperation.End)

    def process_data(self):
        self.filename = self.textbox_filename.toPlainText()
        o_dir = self.textbox_out.toPlainText()
        i_dir = self.textbox_img.toPlainText()
        o_dir_len = len(o_dir)
        i_dir_len = len(i_dir)
        dir_check = True

        if o_dir_len <= 1 or not os.path.isdir(o_dir):
            dir_check = False
            self.textbox_out.setStyleSheet("""
                                                            QTextEdit {
                                                                background-color: rgba(255,255,255,100%);
                                                                font-size : 12px;
                                                                color : red;
                                                                border : 1px Solid Black;
                                                                border-radius : 10px;
                                                                alignment : center;
                                                            }
                                                        """)
            self.textbox_out.setText('Error: Select Directories')
        else:
            print("Good output dir")
        if i_dir_len <= 1 or not os.path.isdir(i_dir):
            dir_check = False
            self.textbox_img.setStyleSheet("""
                                            QTextEdit {
                                                background-color: rgba(255,255,255,100%);
                                                font-size : 12px;
                                                color : red;
                                                border : 1px Solid Black;
                                                border-radius : 10px;
                                                alignment : center;
                                            }
                                            """)
            self.textbox_img.setText('Error: Select Directories')
        else:
            print("Good input dir")
        if dir_check:
            output_dir_list = os.listdir(self.output_dir)
            unique_filename = True

            for nof in output_dir_list:
                if nof == (self.filename + ".xlsx"):
                    unique_filename = False

            if not unique_filename:
                self.textbox_filename_error.setFontPointSize(10)
                self.textbox_filename_error.setText('Error: File name taken!')
            elif self.filename == '':
                self.textbox_filename_error.setFontPointSize(10)
                self.textbox_filename_error.setText('Required!')
            else:
                self.textbox_filename_error.setText('')
                print('Creating output')
                self.format_output()
                self.textbox_status.setFontPointSize(10)
                self.textbox_status.setText('Please Wait...')
                print("Converting images to text...")
                self.convert_image_to_text()
                self.textbox_status.setFontPointSize(10)
                self.textbox_status.setText('Creating Sheet...')
                print("Filling out excel worksheet...")
                self.format_workbook()
                self.textbox_status.setFontPointSize(10)
                self.textbox_status.setText('Finished!')
                print("Done! Please close this window.")

    # Determines the handwritten text from a given image
    def convert_image_to_text(self):
        current_serial_image_path = "N/A"
        current_model_image_path = "N/A"
        serial_value = "N/A"
        img_dir_list = []
        img_temp_list = []
        index = 0

        # populate file names here

        img_dir_list = os.listdir(self.image_dir)
        sort_directory_by_int(img_dir_list)
        print(img_dir_list)
        self.image_dir_size = len(img_dir_list)
        print(str(self.image_dir_size)+' image files found.')

        # sort the list by integers
        img_temp_list.sort(key=int)

        # each pair of images correlates to one item so the loop is indexed by 2
        self.image_list.clear()

        for index in range(0, self.image_dir_size, 2):
            # if the folder of images is formatted as model-image
            m_name = s_name = ''
            if self.orientation_case == self.Orientation.ModelSerial:
                m_name = img_dir_list[index]
                s_name = img_dir_list[index + 1]

            # if the folder of images is formatted as image-model
            elif self.orientation_case == self.Orientation.SerialModel:
                m_name = img_dir_list[index + 1]
                s_name = img_dir_list[index]

            current_model_image_path = self.image_dir + '/' + m_name
            current_serial_image_path = self.image_dir + '/' + s_name
            # google tesseract image to text api
            image = Image.open(current_serial_image_path)

            serial_value = pytesseract.image_to_data(image,
                                                     output_type=pytesseract.Output.DICT,
                                                     lang='eng',
                                                     config="""
                                                            --oem 1
                                                            --psm 7
                                                            -c tessedit_char_whitelist='01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz/\\'
                                                            """)

            # tesseract failed to find a value in the image
            if len(serial_value) == 0:
                serial_value = "Error: Failed to extract value."

            confidence = 0
            text = "Error"

            # append processed row to end of list
            for i in range(len(serial_value['text'])):
                if(serial_value['conf'][i] >= 0):
                    confidence = serial_value['conf'][i]
                    text = serial_value['text'][i]
                    if confidence <= self.confidenceSelected:
                        text = "ERROR: assumption = " + text
                    print(text)
                    print(confidence)

            shutil.copy2(current_model_image_path, self.new_img_dir + "/" + m_name)
            shutil.copy2(current_serial_image_path, self.new_img_dir + "/" + s_name)

            self.image_list.append([self.new_img_dir + "/" + m_name, self.new_img_dir + "/" + s_name, text])
        print('Text extracted successfully.')
    # populates 2xn list with serial&image path

    # Opens user manual in a popup window
    def open_manual(self):
        txtOut = "User manual opened"
        proc = subprocess.Popen(["manual.pdf"], shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

        if(proc.wait() != 0):
            txtOut = "Error: './manual.pdf' failed to open!"

        print(txtOut)

    def orientation_sel(self, orientation):
        if orientation == self.Orientation.ModelSerial:
            print("Orientation: Model Serial")
            self.orientationSelected = self.Orientation.ModelSerial
            self.orientation_button_MS.setStyleSheet("""
                                                     QPushButton {
                                                         background-color : rgba(255,255,255,100%);
                                                         font-size : 12px;
                                                         color: rgba(0,0,0,0%); 
                                                         border : 5px solid blue;
                                                         border-radius : 12px;
                                                         alignment : center;
                                                     }
                                                     QPushButton:hover {
                                                         background-color : rgba(255, 255, 255,100%);
                                                         border : 3px solid blue; 
                                                         border-radius : 12px;
                                                     }
                                                     """)
            self.orientation_button_SM.setStyleSheet("""
                                                     QPushButton {
                                                         background-color : rgba(255,255,255,100%);
                                                         font-size : 12px;
                                                         color: rgba(0,0,0,0%); 
                                                         border : 1px solid black;
                                                         border-radius : 12px;
                                                         alignment : center;
                                                     }
                                                     QPushButton:hover {
                                                         background-color : rgba(255, 255, 255,100%);
                                                         border : 3px solid black; 
                                                         border-radius : 12px;
                                                     }
                                                     """)
        else:
            print("Orientation: Serial Model")
            self.orientationSelected = self.Orientation.SerialModel
            self.orientation_button_SM.setStyleSheet("""
                                                     QPushButton {
                                                         background-color : rgba(255,255,255,100%);
                                                         font-size : 12px;
                                                         color: rgba(0,0,0,0%); 
                                                         border : 5px solid blue;
                                                         border-radius : 12px;
                                                         alignment : center;
                                                     }
                                                     QPushButton:hover {
                                                         background-color : rgba(255, 255, 255,100%);
                                                         border : 3px solid blue; 
                                                         border-radius : 12px;
                                                     }
                                                     """)
            self.orientation_button_MS.setStyleSheet("""
                                                     QPushButton {
                                                         background-color : rgba(255,255,255,100%);
                                                         font-size : 12px;
                                                         color: rgba(0,0,0,0%); 
                                                         border : 1px solid black;
                                                         border-radius : 12px;
                                                         alignment : center;
                                                     }
                                                     QPushButton:hover {
                                                         background-color : rgba(255, 255, 255,100%);
                                                         border : 3px solid black; 
                                                         border-radius : 12px;
                                                     }
                                                     """)

    # Creates output directory
    def format_output(self):
        self.output_dir = self.output_dir + "/" + self.filename
        self.new_img_dir = self.output_dir+"/Data"
        print(self.output_dir)
        print(self.new_img_dir)
        os.mkdir(self.output_dir, 0o777)
        os.mkdir(self.new_img_dir, 0o777)


    # Formats output to excel sheet
    def format_workbook(self):
        self.excel_wb = Workbook()
        self.excel_ws = self.excel_wb.active
        index = 0
        current_row = []
        col_a = "A"
        col_b = "B"
        col_c = "C"
        # Each row needs three columns
        # link for model image
        # link for serial image
        # serial value
        # excel command for local image

        self.excel_ws.title = 'Dataset'
        self.excel_ws.append(['Serial Value', 'Model Image', 'Serial Image'])
        self.excel_ws.row_dimensions[1].height = 24

        self.excel_ws.column_dimensions['A'].width = 71
        self.excel_ws.column_dimensions['B'].width = 71
        self.excel_ws.column_dimensions['C'].width = 71

        self.excel_ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        self.excel_ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
        self.excel_ws['C1'].alignment = Alignment(horizontal='center', vertical='center')

        print('Excel worksheet template created')

        index_length = self.image_dir_size // 2

        print(index_length)
        for index in range(0, index_length, 1):
            print('Reading data entry '+str(index)+':')
            current_row.append(self.image_list[index][2])
            current_row.append('Model Image')
            current_row.append('Serial Image')
            self.excel_ws.append(current_row)
            print('Data entry '+str(index)+' successfully read')

            col_a = "A" + str(index+2)
            col_b = "B" + str(index+2)
            col_c = "C" + str(index+2)

            self.excel_ws[col_a].alignment = Alignment(horizontal='center', vertical='center')

            self.excel_ws[col_b].hyperlink = 'file:///' + self.image_list[index][0]
            self.excel_ws[col_b].style = 'Hyperlink'
            self.excel_ws[col_b].alignment = Alignment(horizontal='center', vertical='center')

            self.excel_ws[col_c].hyperlink = 'file:///' + self.image_list[index][1]
            self.excel_ws[col_c].style = 'Hyperlink'
            self.excel_ws[col_c].alignment = Alignment(horizontal='center', vertical='center')

            current_row.clear()

        print('Data successfully exported to excel sheet')

        self.excel_wb.save(self.output_dir + "/Dataset.xlsx")
        print('Excel sheet successfully created')

if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = Window()
    window.show()

    try:
        sys.exit(app.exec())
    except SystemExit:
        print('Shutting Down')


