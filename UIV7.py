import sys
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6 import QtCore
from PyQt6.QtCore import Qt
import random
import time
from statistics import stdev
from statistics import variance
from statistics import mean
import pandas as pd
import numpy as np
import serial
from pandas import option_context
from PyQt6 import QtWidgets
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtSerialPort import QSerialPort
import openpyxl
import serial.tools.list_ports
# voltage = True
# arduino = arduinolib.Device()
pg1 = False
started = False

buttonstyleoff = """
            QPushButton {
                border-radius: 10px;
                padding: 20px;
                background-color: #DFE3E3;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
                border-radius: 10px;
            }
            """
buttonstyleon = """
            QPushButton {
                color = red;
                border-radius: 10px;
                padding: 20px;
                background-color: #D7AB29;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #D7AB29;
                border-radius: 10px;
            }
            """
progessstyle = """
    QProgressBar {
        border: 2px solid rgba(33, 37, 43, 180);
        border-radius: 5px;
        text-align: center;
        background-color: #DFE3E3;
        color: black;
    }
    QProgressBar::chunk {
        background-color: #D7AB29;
    }
"""

def detect_arduino_port():
    arduino_ports = [
        p.device
        for p in serial.tools.list_ports.comports()
        if 'Arduino' in p.description  # Modify this condition to match your Arduino's description
    ]
    
    if not arduino_ports:
        raise IOError("No Arduino found")
    
    # If multiple Arduinos are connected, you can choose one based on your preference
    # For example, you can return the first Arduino port from the list
    return arduino_ports[0]

# Usage
try:
    arduino_port = detect_arduino_port()
except IOError as e:
    app = QApplication(sys.argv)
    message_box = QMessageBox()
    message_box.setWindowTitle("Arduino Connection Error")
    message_box.setIcon(QMessageBox.Icon.Critical)
    message_box.setText("No Arduino device detected. Please connect an Arduino and reopen the application.")
    message_box.exec()
    sys.exit()

class SerialThread(QThread):
    received = pyqtSignal(str)

    def __init__(self, port):
        super().__init__()
        self.serial_port = serial.Serial(port, 115200, timeout=1)

    def run(self):
        while True:
            if self.serial_port.in_waiting:
                serial_data = self.serial_port.readline().decode(errors='ignore').strip()
                self.received.emit(serial_data)

    def write_data(self, data):
        self.serial_port.write(data)
    def stop(self):
        self.running = False
    def clear_buffer(self):
        self.serial_port.reset_input_buffer()
        self.serial_port.reset_output_buffer()

# serial_thread = SerialThread('COM20')
class Page1(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.blanklabel = QLabel("")        
        # label1 = QLabel(self)
        label2 = QLabel(self)
        # pixmap = QPixmap('PAFDCLogo.png')
        # label1.setPixmap(pixmap)
        pixmap2 = QPixmap('Multi_touch.png')
        label2.setPixmap(pixmap2)
        label2.setMaximumSize(600,200)
        page1layout = QGridLayout()
        # page1layout.addWidget(label1,2,2)
        page1layout.addWidget(label2,1,1)
        page1layout.addWidget(self.blanklabel,0,0)
        page1layout.addWidget(self.blanklabel,0,1)
        page1layout.addWidget(self.blanklabel,0,2)
        page1layout.addWidget(self.blanklabel,1,0)
        page1layout.addWidget(self.blanklabel,2,0)
        # page1layout.addWidget(self.blanklabel,3,0)
        # page1layout.addWidget(self.blanklabel,3,0)
        # page1layout.addWidget(self.blanklabel,4,0)
        self.setLayout(page1layout)
class Page2(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.label = QLabel("Voltmeter")
        self.label2 = QLabel("")
        self.page2layout = QGridLayout()
        self.blanklabel = QLabel("")
        self.label3 = QLabel("")
        self.label5= QLabel("")
        self.startbutton = QPushButton("Start/Reset")
        self.stopbutton = QPushButton("Stop")
        self.tentrialsb = QPushButton("10 Trials")
        self.fiftytrialsb = QPushButton("50 Trials")
        self.hundredtrialsb = QPushButton("100 Trials")
        self.csvbutton = QPushButton("Save to Excel")
        self.dontsave = QPushButton("Don't Save")
        self.dontsave.setFixedSize(127,57)
        self.dontsave.setStyleSheet(buttonstyleoff)
        self.dontsave.hide()
        self.textbox = QLineEdit(self)
        self.textbox.setFixedSize(127,20)
  
        self.label4 = QLabel("                                                     Filename:")
        self.label4.hide()
        self.textbox.hide()
        self.textbox.returnPressed.connect(self.csvbutton.click)
        self.startbutton.setFixedSize(127,57)
        self.stopbutton.setFixedSize(127,57)
        self.tentrialsb.setFixedSize(127,57)
        self.fiftytrialsb.setFixedSize(127,57)
        self.hundredtrialsb.setFixedSize(127,57)
        self.csvbutton.setFixedSize(127,57)
        self.startbutton.setStyleSheet(buttonstyleoff)
        self.stopbutton.setStyleSheet(buttonstyleon) 
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.csvbutton.hide()
        self.page2layout.addWidget(self.textbox,2,1)
        self.page2layout.addWidget(self.label5,4,1)
        self.page2layout.addWidget(self.label4,2,0)
        self.page2layout.addWidget(self.csvbutton,3,1)
        self.page2layout.addWidget(self.blanklabel,0,1)
        self.page2layout.addWidget(self.blanklabel,0,2)
        self.page2layout.addWidget(self.blanklabel,0,3)
        self.page2layout.addWidget(self.blanklabel,0,4)
        self.page2layout.addWidget(self.blanklabel,1,0)
        self.page2layout.addWidget(self.blanklabel,2,0)
        self.page2layout.addWidget(self.blanklabel,3,0)
        self.page2layout.addWidget(self.blanklabel,4,0)
        self.page2layout.addWidget(self.startbutton,4,3)
        self.page2layout.addWidget(self.stopbutton,4,4)
        self.page2layout.addWidget(self.label,0,0)
        self.page2layout.addWidget(self.label2,1,2)
        self.page2layout.addWidget(self.tentrialsb,2,0)
        self.page2layout.addWidget(self.fiftytrialsb,3,0)
        self.page2layout.addWidget(self.hundredtrialsb,4,0)
        self.page2layout.addWidget(self.label3,0,4)
        self.page2layout.addWidget(self.dontsave,4,1)


        self.startbutton.clicked.connect(self.start)
        self.stopbutton.clicked.connect(self.stop)
        self.tentrialsb.clicked.connect(self.tentrials)
        self.fiftytrialsb.clicked.connect(self.fiftytrials)
        self.hundredtrialsb.clicked.connect(self.hundredtrials)
        fontheader = QFont("Arial", 40)
        self.label.setFont(fontheader)
        font = QFont("Arial", 30)
        self.label2.setFont(font) 
        self.setLayout(self.page2layout)
        self.count = 0
        self.timer = QTimer()
        # self.timer.timeout.connect(self.updateCount)
        self.timer.start(50)
        # self.serial_thread = SerialThread("COM14")  # Adjust the port name as per your system
        # self.serial_thread.received.connect(self.print_serial_data)
        # self.serial_thread.start()
        # def print_serial_data(self, data):
        # print(data)    

    def namesave(self,x):
        name = self.textbox.text()
        sd = stdev(x)
        var = variance(x)
        m = mean(x)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label4.isVisible():
            self.label4.hide()
        if name == "":
            timestr = time.strftime("%Y%m%d_%H%M%S")
            name = "{}.xlsx".format(timestr)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        else:
            name = "{}.xlsx".format(name)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        col = list(range(1, len(x) + 1))
        df = pd.DataFrame({'Trial #': col, 'Voltage': x})
        df.at[0, 'Std Dev.'] = '{:.2f}'.format(sd)
        df.at[0, 'Variance'] = '{:.2f}'.format(var)
        df.at[0, 'Mean'] = '{:.2f}'.format(m)
        with pd.ExcelWriter('{}'.format(name)) as writer:
            df.style.set_properties(**{'text-align': 'center'}).to_excel(writer,sheet_name='Sheet1',index = False)
        self.textbox.clear()
        serial_thread.wait()
        return


    def updatemeter(self):
        serial_thread.received.connect(self.print_serial_data)
            
    def print_serial_data(self, data):
        self.label2.setText(f"{data}")

    def start(self):
        serial_thread.write_data(b'1')
        serial_thread.start()
        self.label3.setText("")
        self.label5.setText("")
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.textbox.hide()
        self.label4.hide()
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.timer.timeout.connect(lambda: self.updatemeter())
        
    def stop(self):
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.label2.setText("")
        serial_thread.terminate()
        # self.a = []

    def tentrials(self):
        self.stop()
        serial_thread.write_data(b'1')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleon)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.csvbutton = QPushButton("Save to Excel")
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page2layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page2layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.a = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_10_trials())
        # for i in range(len(ten)):
        #     x = random.randint(0, 9)
        #     self.a.append(x)
        #     time.sleep(.05)
        # self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a),variance(self.a),mean(self.a)))
        # # self.tentrialsb.setStyleSheet(buttonStyle1)
        # # self.stopbutton.setStyleSheet(buttonStyle1)
        # # self.startbutton.setStyleSheet(buttonStyle)
        # self.csvbutton.clicked.connect(lambda: [self.namesave(self.a),self.csvbutton.deleteLater()])
        # self.tentrialsb.destroy()
        # self.stop()
    def savereset(self):
        self.csvbutton.hide()
        self.label4.hide()
        self.dontsave.hide()
        self.label2.setText("")
        self.label5.setText("")
        self.label3.setText("")
        self.startbutton.setEnabled(True)
        self.stopbutton.setEnabled(True)
        self.tentrialsb.setEnabled(True)
        self.fiftytrialsb.setEnabled(True)
        self.hundredtrialsb.setEnabled(True)
        self.textbox.hide()
        self.stop()

    def update_meter_10_trials(self):
        self.stop()
        if len(self.a) < 10:
            self.progress.show()
            self.progress.setValue((len(self.a)+1)*10)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                else:
                    self.a.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a), variance(self.a), mean(self.a)))
            self.stop()
            self.timer.timeout.disconnect()
            self.csvbutton.show()
            self.textbox.show()
            self.label4.show()
            self.dontsave.show()
            self.dontsave.clicked.connect(lambda: self.savereset())
            self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
            self.tentrialsb.destroy()

    def fiftytrials(self):
        self.stop()
        serial_thread.write_data(b'1')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleon)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page2layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page2layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.b = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_50_trials())

    def update_meter_50_trials(self):
        self.stop()
        if len(self.b) < 50:
            self.progress.show()
            self.progress.setValue((len(self.b)+1)*2)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                else:
                    self.b.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.b), variance(self.b), mean(self.b)))
            self.timer.timeout.disconnect()
            self.csvbutton.show()
            self.textbox.show()
            self.label4.show()
            self.dontsave.show()
            self.dontsave.clicked.connect(lambda: self.savereset())
            self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
            self.fiftytrialsb.destroy()

    def hundredtrials(self):
        self.stop()
        serial_thread.write_data(b'1')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleon)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page2layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page2layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.c = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_100_trials())
    def update_meter_100_trials(self):
        self.stop()
        if len(self.c) < 100:
            self.progress.show()
            self.progress.setValue((len(self.c)+1))
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                else:
                    self.c.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.c), variance(self.c), mean(self.c)))
            self.stop()
            self.timer.timeout.disconnect()
            self.csvbutton.show()
            self.textbox.show()
            self.label4.show()
            self.dontsave.show()
            self.dontsave.clicked.connect(lambda: self.savereset())
            self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
            self.tentrialsb.destroy()

    def hideEvent(self, event):
        self.stop()
class Page3(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.label = QLabel("Diode/LED Test")
        buttonStyle = """
            QPushButton {
                border-radius: 10px;
                padding: 20px;
                background-color: #DFE3E3;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
                border-radius: 10px;
            }
            """
        buttonStyle1 = """
            QPushButton {
                border-radius: 10px;
                padding: 20px;
                background-color: #D7AB29;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
                border-radius: 10px;
            }
            """
        self.label2 = QLabel("")
        self.page3layout = QGridLayout()
        self.blanklabel = QLabel("")
        self.label5= QLabel("")
        self.label3 = QLabel("")
        self.startbutton = QPushButton("Start/Reset")
        self.stopbutton = QPushButton("Stop")
        self.tentrialsb = QPushButton("10 Trials")
        self.fiftytrialsb = QPushButton("50 Trials")
        self.hundredtrialsb = QPushButton("100 Trials")
        self.csvbutton = QPushButton("Save to Excel")
        self.textbox = QLineEdit(self)
        self.textbox.setFixedSize(127,20)
        self.label4 = QLabel("                                                     Filename:")
        self.label4.hide()
        self.textbox.hide()
        self.startbutton.setFixedSize(127,57)
        self.stopbutton.setFixedSize(127,57)
        self.tentrialsb.setFixedSize(127,57)
        self.fiftytrialsb.setFixedSize(127,57)
        self.hundredtrialsb.setFixedSize(127,57)
        self.csvbutton.setFixedSize(127,57)
        self.textbox.returnPressed.connect(self.csvbutton.click)
        self.startbutton.setStyleSheet(buttonstyleoff)
        self.stopbutton.setStyleSheet(buttonstyleon)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.csvbutton.hide()
        self.dontsave = QPushButton("Don't Save")
        self.dontsave.setFixedSize(127,57)
        self.dontsave.setStyleSheet(buttonStyle)
        
        self.dontsave.hide()
        self.page3layout.addWidget(self.textbox,2,1)
        self.page3layout.addWidget(self.label5,4,1)
        self.page3layout.addWidget(self.label4,2,0)
        self.page3layout.addWidget(self.csvbutton,3,1)
        self.page3layout.addWidget(self.blanklabel,0,1)
        self.page3layout.addWidget(self.blanklabel,0,2)
        self.page3layout.addWidget(self.blanklabel,0,3)
        self.page3layout.addWidget(self.blanklabel,0,4)
        self.page3layout.addWidget(self.blanklabel,1,0)
        self.page3layout.addWidget(self.blanklabel,2,0)
        self.page3layout.addWidget(self.blanklabel,3,0)
        self.page3layout.addWidget(self.blanklabel,4,0)
        self.page3layout.addWidget(self.label2,1,2)
        self.page3layout.addWidget(self.startbutton,4,3)
        self.page3layout.addWidget(self.stopbutton,4,4)
        self.page3layout.addWidget(self.tentrialsb,2,0)
        self.page3layout.addWidget(self.fiftytrialsb,3,0)
        self.page3layout.addWidget(self.hundredtrialsb,4,0)
        self.page3layout.addWidget(self.label3,0,4)
        self.page3layout.addWidget(self.dontsave,4,1)
        self.startbutton.clicked.connect(self.start)
        self.stopbutton.clicked.connect(self.stop)
        self.tentrialsb.clicked.connect(self.tentrials)
        self.fiftytrialsb.clicked.connect(self.fiftytrials)
        self.hundredtrialsb.clicked.connect(self.hundredtrials)
        self.page3layout.addWidget(self.label,0,0)
        fontheader = QFont("Arial", 40)
        self.label.setFont(fontheader)
        font = QFont("Arial", 30)
        self.label2.setFont(font) 
        self.setLayout(self.page3layout)
        self.timer = QTimer()
        self.count = 0
        # self.timer.timeout.connect(self.updateCount)
        self.timer.start(50)

    def updatemeter(self):
        serial_thread.received.connect(self.print_serial_data)
            
    def print_serial_data(self, data):
        self.label2.setText(f"{data}")

    def namesave(self,x):
        name = self.textbox.text()
        sd = stdev(x)
        var = variance(x)
        m = mean(x)
        
        if name == "":
            timestr = time.strftime("%Y%m%d_%H%M%S")
            name = "{}.xlsx".format(timestr)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        else:
            name = "{}.xlsx".format(name)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        col = list(range(1, len(x) + 1))
        df = pd.DataFrame({'Trial #': col, 'Voltage': x})
        df.at[0, 'Std Dev.'] = '{:.2f}'.format(sd)
        df.at[0, 'Variance'] = '{:.2f}'.format(var)
        df.at[0, 'Mean'] = '{:.2f}'.format(m)
        with pd.ExcelWriter('{}'.format(name)) as writer:
            df.style.set_properties(**{'text-align': 'center'}).to_excel(writer,sheet_name='Sheet1',index = False)
        self.textbox.clear()
        serial_thread.wait()
        self.stop()
        return
    def start(self):
        serial_thread.write_data(b'2')
        serial_thread.start()
        self.label3.setText("")
        self.label5.setText("")
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.textbox.hide()
        self.label4.hide()
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.timer.timeout.connect(lambda: self.updatemeter())

    def stop(self):
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.label2.setText("")
        serial_thread.terminate()
        # self.a = []
    def savereset(self):
        self.csvbutton.hide()
        self.label4.hide()
        self.dontsave.hide()
        self.label2.setText("")
        self.label3.setText("")
        self.label5.setText("")
        self.startbutton.setEnabled(True)
        self.stopbutton.setEnabled(True)
        self.tentrialsb.setEnabled(True)
        self.fiftytrialsb.setEnabled(True)
        self.hundredtrialsb.setEnabled(True)
        self.textbox.hide()
        self.stop()


    def tentrials(self):
        self.stop()
        serial_thread.write_data(b'2')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleon)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page3layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page3layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.a = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_10_trials())
        # for i in range(len(ten)):
        #     x = random.randint(0, 9)
        #     self.a.append(x)
        #     time.sleep(.05)
        # self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a),variance(self.a),mean(self.a)))
        # # self.tentrialsb.setStyleSheet(buttonStyle1)
        # # self.stopbutton.setStyleSheet(buttonStyle1)
        # # self.startbutton.setStyleSheet(buttonStyle)
        # self.csvbutton.clicked.connect(lambda: [self.namesave(self.a),self.csvbutton.deleteLater()])
        # self.tentrialsb.destroy()
        # self.stop()
    def update_meter_10_trials(self):
        self.stop()
        if len(self.a) < 10:
            self.progress.show()
            self.progress.setValue((len(self.a)+1)*10)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                else:
                    self.a.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a), variance(self.a), mean(self.a)))
            self.stop()
            self.timer.timeout.disconnect()
            self.csvbutton.show()
            self.textbox.show()
            self.label4.show()
            self.dontsave.show()
            self.dontsave.clicked.connect(lambda: self.savereset())
            self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
            self.tentrialsb.destroy()


    def fiftytrials(self):
        self.stop()
        serial_thread.write_data(b'2')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleon)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page3layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page3layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.b = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_50_trials())

    def update_meter_50_trials(self):
        self.stop()
        if len(self.b) < 50:
            self.progress.show()
            self.progress.setValue((len(self.b)+1)*2)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                else:
                    self.b.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.b), variance(self.b), mean(self.b)))
            self.timer.timeout.disconnect()
            self.csvbutton.show()
            self.textbox.show()
            self.label4.show()
            self.dontsave.show()
            self.dontsave.clicked.connect(lambda: self.savereset())
            self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])

    def hundredtrials(self):
        self.stop()
        serial_thread.write_data(b'2')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleon)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page3layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page3layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.c = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_100_trials())
    def update_meter_100_trials(self):
        self.stop()
        if len(self.c) < 100:
            self.progress.show()
            self.progress.setValue((len(self.c)+1))
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                else:
                    self.c.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.c), variance(self.c), mean(self.c)))
            self.stop()
            self.timer.timeout.disconnect()
            self.csvbutton.show()
            self.textbox.show()
            self.label4.show()
            self.dontsave.show()
            self.dontsave.clicked.connect(lambda: self.savereset())
            self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
            self.tentrialsb.destroy()
    def hideEvent(self, event):
        self.stop()
class Page4(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        buttonStyle = """
            QPushButton {
                border-radius: 10px;
                padding: 20px;
                background-color: #DFE3E3;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
                border-radius: 10px;
            }
            """
        buttonStyle1 = """
            QPushButton {
                border-radius: 10px;
                padding: 20px;
                background-color: #D7AB29;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
                border-radius: 10px;
            }
            """

        label = QLabel("Ohmmeter")
        self.label2 = QLabel("")
        font = QFont("Arial", 30)
        self.label2.setFont(font) 
        fontheader = QFont("Arial", 40)
        label.setFont(fontheader)
        self.page4layout = QGridLayout()
        self.label5= QLabel("")
        self.label3 = QLabel("")
        self.blanklabel = QLabel("")
        self.startbutton = QPushButton("Start/Reset")
        self.stopbutton = QPushButton("Stop")
        self.tentrialsb = QPushButton("10 Trials")
        self.fiftytrialsb = QPushButton("50 Trials")
        self.hundredtrialsb = QPushButton("100 Trials")
        self.csvbutton = QPushButton("Save to Excel")
        self.textbox = QLineEdit(self)
        self.textbox.setFixedSize(127,20)
        self.label4 = QLabel("                                                     Filename:")
        self.label4.hide()
        self.textbox.hide()
        self.startbutton.setFixedSize(127,57)
        self.stopbutton.setFixedSize(127,57)
        self.tentrialsb.setFixedSize(127,57)
        self.fiftytrialsb.setFixedSize(127,57)
        self.hundredtrialsb.setFixedSize(127,57)
        self.csvbutton.setFixedSize(127,57)
        self.startbutton.setStyleSheet(buttonstyleoff)
        self.stopbutton.setStyleSheet(buttonstyleon)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.csvbutton.hide()
        self.dontsave = QPushButton("Don't Save")
        self.dontsave.setFixedSize(127,57)
        self.dontsave.setStyleSheet(buttonStyle)
        
        self.dontsave.hide()
        self.page4layout.addWidget(self.textbox,2,1)
        self.page4layout.addWidget(self.label5,4,1)
        self.page4layout.addWidget(self.label4,2,0)
        self.page4layout.addWidget(self.csvbutton,3,1)
        self.page4layout.addWidget(self.blanklabel,0,1)
        self.page4layout.addWidget(self.blanklabel,0,2)
        self.page4layout.addWidget(self.blanklabel,0,3)
        self.page4layout.addWidget(self.blanklabel,0,4)
        self.page4layout.addWidget(self.blanklabel,1,0)
        self.page4layout.addWidget(self.blanklabel,2,0)
        self.page4layout.addWidget(self.blanklabel,3,0)
        self.page4layout.addWidget(self.blanklabel,4,0)
        self.page4layout.addWidget(self.label2,1,2)
        self.page4layout.addWidget(self.startbutton,4,3)
        self.page4layout.addWidget(self.stopbutton,4,4)
        self.page4layout.addWidget(self.tentrialsb,2,0)
        self.page4layout.addWidget(self.fiftytrialsb,3,0)
        self.page4layout.addWidget(self.hundredtrialsb,4,0)
        self.page4layout.addWidget(self.label3,0,4)
        self.page4layout.addWidget(self.dontsave,4,1)
        self.startbutton.clicked.connect(self.start)
        self.stopbutton.clicked.connect(self.stop)
        self.tentrialsb.clicked.connect(self.tentrials)
        self.fiftytrialsb.clicked.connect(self.fiftytrials)
        self.hundredtrialsb.clicked.connect(self.hundredtrials)
        self.page4layout.addWidget(label,0,0)
        self.setLayout(self.page4layout)
        self.timer = QTimer()
        self.count = 0
        # self.timer.timeout.connect(self.updateCount)
        self.timer.start(50)
    def namesave(self,x):
        name = self.textbox.text()
        sd = stdev(x)
        var = variance(x)
        m = mean(x)
        
        if name == "":
            timestr = time.strftime("%Y%m%d_%H%M%S")
            name = "{}.xlsx".format(timestr)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        else:
            name = "{}.xlsx".format(name)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        col = list(range(1, len(x) + 1))
        df = pd.DataFrame({'Trial #': col, 'Resistance': x})
        df.at[0, 'Std Dev.'] = '{:.2f}'.format(sd)
        df.at[0, 'Variance'] = '{:.2f}'.format(var)
        df.at[0, 'Mean'] = '{:.2f}'.format(m)
        with pd.ExcelWriter('{}'.format(name)) as writer:
            df.style.set_properties(**{'text-align': 'center'}).to_excel(writer,sheet_name='Sheet1',index = False)
        self.textbox.clear()
        serial_thread.wait()
        self.stop()
        return
    def updatemeter(self):
        serial_thread.received.connect(self.print_serial_data)
            
    def print_serial_data(self, data):
        self.label2.setText(f"{data}")
    def savereset(self):
        self.csvbutton.hide()
        self.label4.hide()
        self.dontsave.hide()
        self.label2.setText("")
        self.label3.setText("")
        self.label5.setText("")
        self.startbutton.setEnabled(True)
        self.stopbutton.setEnabled(True)
        self.tentrialsb.setEnabled(True)
        self.fiftytrialsb.setEnabled(True)
        self.hundredtrialsb.setEnabled(True)
        self.textbox.hide()
        self.stop()

    def start(self):
        serial_thread.write_data(b'3')
        serial_thread.start()
        self.label3.setText("")
        self.label5.setText("")
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.textbox.hide()
        self.label4.hide()
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.timer.timeout.connect(lambda: self.updatemeter())
    def stop(self):
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.label2.setText("")
        serial_thread.terminate()
    def tentrials(self):
        self.stop()
        serial_thread.write_data(b'3')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleon)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page4layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page4layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.a = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_10_trials())
        # for i in range(len(ten)):
        #     x = random.randint(0, 9)
        #     self.a.append(x)
        #     time.sleep(.05)
        # self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a),variance(self.a),mean(self.a)))
        # # self.tentrialsb.setStyleSheet(buttonStyle1)
        # # self.stopbutton.setStyleSheet(buttonStyle1)
        # # self.startbutton.setStyleSheet(buttonStyle)
        # self.csvbutton.clicked.connect(lambda: [self.namesave(self.a),self.csvbutton.deleteLater()])
        # self.tentrialsb.destroy()
        # self.stop()
    def update_meter_10_trials(self):
        self.stop()
        if len(self.a) < 10:
            self.progress.show()
            self.progress.setValue((len(self.a)+1)*10)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                elif serial_data == "inf":
                    self.a.append(float('inf'))
                else:
                    self.a.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            f = sum(self.a)
            f = str(f)
            if f == 'inf':
                self.label3.setText("Std Deviation: {}\nVariance: {}\nMean: {}".format('inf', 'inf', 'inf'))
                self.timer.timeout.disconnect()
                self.stop()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
            else:
                self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a), variance(self.a), mean(self.a)))
                self.stop()
                self.timer.timeout.disconnect()
                self.csvbutton.show()
                self.textbox.show()
                self.label4.show()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
                self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
                self.tentrialsb.destroy()

    def fiftytrials(self):
        self.stop()
        serial_thread.write_data(b'3')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleon)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page4layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page4layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.b = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_50_trials())

    def update_meter_50_trials(self):
        self.stop()
        if len(self.b) < 50:
            self.progress.show()
            self.progress.setValue((len(self.b)+1)*2)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                elif serial_data == "inf":
                    self.b.append(float('inf'))
                else:
                    self.b.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            f = sum(self.b)
            f = str(f)
            if f == 'inf':
                self.label3.setText("Std Deviation: {}\nVariance: {}\nMean: {}".format('inf', 'inf', 'inf'))
                self.timer.timeout.disconnect()
                self.stop()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
            else:
                self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.b), variance(self.b), mean(self.b)))
                self.stop()
                self.timer.timeout.disconnect()
                self.csvbutton.show()
                self.textbox.show()
                self.label4.show()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
                self.csvbutton.clicked.connect(lambda: [self.namesave(self.b), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
                self.tentrialsb.destroy()

    def hundredtrials(self):
        self.stop()
        serial_thread.write_data(b'3')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleon)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page4layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page4layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.c = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_100_trials())
    def update_meter_100_trials(self):
        self.stop()
        if len(self.c) < 100:
            self.progress.show()
            self.progress.setValue((len(self.c)+1))
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                elif serial_data == "inf":
                    self.c.append(float('inf'))
                else:
                    self.c.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            f = sum(self.c)
            f = str(f)
            if f == 'inf':
                self.label3.setText("Std Deviation: {}\nVariance: {}\nMean: {}".format('inf', 'inf', 'inf'))
                self.timer.timeout.disconnect()
                self.stop()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
            else:
                self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.c), variance(self.c), mean(self.c)))
                self.stop()
                self.timer.timeout.disconnect()
                self.csvbutton.show()
                self.textbox.show()
                self.label4.show()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
                self.csvbutton.clicked.connect(lambda: [self.namesave(self.c), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
                self.tentrialsb.destroy()

    def hideEvent(self, event):
        self.stop()
class Page5(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        buttonStyle = """
            QPushButton {
                border-radius: 10px;
                padding: 20px;
                background-color: #DFE3E3;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
                border-radius: 10px;
            }
            """
        buttonStyle1 = """
            QPushButton {
                border-radius: 10px;
                padding: 20px;
                background-color: #D7AB29;
            }
            QPushButton:hover {
                background-color: #c0c0c0;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
                border-radius: 10px;
            }
            """
        self.label = QLabel("Continuity Test")
        fontheader = QFont("Arial", 30)
        self.label.setFont(fontheader)
        self.label2 = QLabel("")
        self.label2.setText("")
        font = QFont("Arial", 30)
        self.label2.setFont(font)
        self.label5= QLabel("")
        self.label3 = QLabel("")
        self.page5layout = QGridLayout()
        self.startbutton = QPushButton("Start/Reset")
        self.stopbutton = QPushButton("Stop")
        self.tentrialsb = QPushButton("10 Trials")
        self.fiftytrialsb = QPushButton("50 Trials")
        self.hundredtrialsb = QPushButton("100 Trials")
        self.csvbutton = QPushButton("Save to Excel")
        self.textbox = QLineEdit(self)
        self.textbox.setFixedSize(127,20)
        self.label4 = QLabel("                                                     Filename:")
        self.label4.hide()
        self.textbox.hide()
        self.startbutton.setFixedSize(127,57)
        self.stopbutton.setFixedSize(127,57)
        self.tentrialsb.setFixedSize(127,57)
        self.fiftytrialsb.setFixedSize(127,57)
        self.hundredtrialsb.setFixedSize(127,57)
        self.csvbutton.setFixedSize(127,57)
        self.textbox.returnPressed.connect(self.csvbutton.click)
        self.startbutton.setStyleSheet(buttonstyleoff)
        self.stopbutton.setStyleSheet(buttonstyleon)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.csvbutton.hide()
        self.blanklabel = QLabel("")
        self.dontsave = QPushButton("Don't Save")
        self.dontsave.setFixedSize(127,57)
        self.dontsave.setStyleSheet(buttonStyle)
        
        self.dontsave.hide()
        self.page5layout.addWidget(self.textbox,2,1)
        self.page5layout.addWidget(self.label5,4,1)
        self.page5layout.addWidget(self.label4,2,0)
        self.page5layout.addWidget(self.csvbutton,3,1)
        self.page5layout.addWidget(self.blanklabel,0,1)
        self.page5layout.addWidget(self.blanklabel,0,2)
        self.page5layout.addWidget(self.blanklabel,0,3)
        self.page5layout.addWidget(self.blanklabel,0,4)
        self.page5layout.addWidget(self.blanklabel,1,0)
        self.page5layout.addWidget(self.blanklabel,2,0)
        self.page5layout.addWidget(self.blanklabel,3,0)
        self.page5layout.addWidget(self.blanklabel,4,0)
        self.page5layout.addWidget(self.label2,1,2)
        self.page5layout.addWidget(self.startbutton,4,3)
        self.page5layout.addWidget(self.stopbutton,4,4)
        self.page5layout.addWidget(self.tentrialsb,2,0)
        self.page5layout.addWidget(self.fiftytrialsb,3,0)
        self.page5layout.addWidget(self.hundredtrialsb,4,0)
        self.page5layout.addWidget(self.label3,0,4)
        self.page5layout.addWidget(self.label,0,0)
        self.page5layout.addWidget(self.dontsave,4,1)
        self.startbutton.clicked.connect(self.start)
        self.stopbutton.clicked.connect(self.stop)
        self.tentrialsb.clicked.connect(self.tentrials)
        self.fiftytrialsb.clicked.connect(self.fiftytrials)
        self.hundredtrialsb.clicked.connect(self.hundredtrials)
        self.setLayout(self.page5layout)
        self.timer = QTimer()
        self.count = 0
        self.timer.start(50)

    def namesave(self,x):
        name = self.textbox.text()
        sd = stdev(x)
        var = variance(x)
        m = mean(x)
        
        if name == "":
            timestr = time.strftime("%Y%m%d_%H%M%S")
            name = "{}.xlsx".format(timestr)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        else:
            name = "{}.xlsx".format(name)
            self.label5.setText("Saved as {}".format(name))
            self.textbox.clear()
        col = list(range(1, len(x) + 1))
        df = pd.DataFrame({'Trial #': col, 'Resistance': x})
        df.at[0, 'Std Dev.'] = '{:.2f}'.format(sd)
        df.at[0, 'Variance'] = '{:.2f}'.format(var)
        df.at[0, 'Mean'] = '{:.2f}'.format(m)
        with pd.ExcelWriter('{}'.format(name)) as writer:
            df.style.set_properties(**{'text-align': 'center'}).to_excel(writer,sheet_name='Sheet1',index = False)
        self.textbox.clear()
        serial_thread.wait()
        self.stop()
        return
    
    def updatemeter(self):
        serial_thread.received.connect(self.print_serial_data)

    def print_serial_data(self, data):
        self.label2.setText(f"{data}")
    def savereset(self):
        self.csvbutton.hide()
        self.label4.hide()
        self.dontsave.hide()
        self.label3.setText("")
        self.label5.setText("")
        self.label2.setText("")
        self.startbutton.setEnabled(True)
        self.stopbutton.setEnabled(True)
        self.tentrialsb.setEnabled(True)
        self.fiftytrialsb.setEnabled(True)
        self.hundredtrialsb.setEnabled(True)
        self.textbox.hide()
        self.stop()

    def start(self):
        serial_thread.write_data(b'4')
        serial_thread.start()
        self.label3.setText("")
        self.label5.setText("")
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.textbox.hide()
        self.label4.hide()
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.timer.timeout.connect(lambda: self.updatemeter())
        
    def stop(self):
        self.startbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.stopbutton.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.label2.setText("")
        serial_thread.terminate()
    def tentrials(self):
        self.stop()
        serial_thread.write_data(b'4')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.tentrialsb.setStyleSheet(buttonstyleon)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page5layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page5layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.a = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_10_trials())
        # for i in range(len(ten)):
        #     x = random.randint(0, 9)
        #     self.a.append(x)
        #     time.sleep(.05)
        # self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a),variance(self.a),mean(self.a)))
        # # self.tentrialsb.setStyleSheet(buttonStyle1)
        # # self.stopbutton.setStyleSheet(buttonStyle1)
        # # self.startbutton.setStyleSheet(buttonStyle)
        # self.csvbutton.clicked.connect(lambda: [self.namesave(self.a),self.csvbutton.deleteLater()])
        # self.tentrialsb.destroy()
        # self.stop()
    def update_meter_10_trials(self):
        self.stop()
        if len(self.a) < 10:
            self.progress.show()
            self.progress.setValue((len(self.a)+1)*10)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                elif serial_data == "inf":
                    self.a.append(float('inf'))
                else:
                    self.a.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            f = sum(self.a)
            f = str(f)
            if f == 'inf':
                self.label3.setText("Std Deviation: {}\nVariance: {}\nMean: {}".format('inf', 'inf', 'inf'))
                self.timer.timeout.disconnect()
                self.stop()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
            else:
                self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.a), variance(self.a), mean(self.a)))
                self.stop()
                self.timer.timeout.disconnect()
                self.csvbutton.show()
                self.textbox.show()
                self.label4.show()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
                self.csvbutton.clicked.connect(lambda: [self.namesave(self.a), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
                self.tentrialsb.destroy()


    def fiftytrials(self):
        self.stop()
        serial_thread.write_data(b'4')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleon)
        self.hundredtrialsb.setStyleSheet(buttonstyleoff)
        self.csvbutton = QPushButton("Save to Excel")
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page5layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page5layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.b = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_50_trials())

    def update_meter_50_trials(self):
        self.stop()
        if len(self.b) < 50:
            self.progress.show()
            self.progress.setValue((len(self.b)+1)*2)
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                elif serial_data == "inf":
                    self.b.append(float('inf'))
                else:
                    self.b.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            f = sum(self.b)
            f = str(f)
            if f == 'inf':
                self.label3.setText("Std Deviation: {}\nVariance: {}\nMean: {}".format('inf', 'inf', 'inf'))
                self.timer.timeout.disconnect()
                self.stop()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
            else:
                self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.b), variance(self.b), mean(self.b)))
                self.stop()
                self.timer.timeout.disconnect()
                self.csvbutton.show()
                self.textbox.show()
                self.label4.show()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
                self.csvbutton.clicked.connect(lambda: [self.namesave(self.b), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
                self.tentrialsb.destroy()

    def hundredtrials(self):
        self.stop()
        serial_thread.write_data(b'4')
        serial_thread.start()
        self.startbutton.setEnabled(False)
        self.startbutton.setEnabled(False)
        self.tentrialsb.setEnabled(False)
        self.fiftytrialsb.setEnabled(False)
        self.hundredtrialsb.setEnabled(False)
        self.tentrialsb.setStyleSheet(buttonstyleoff)
        self.fiftytrialsb.setStyleSheet(buttonstyleoff)
        self.hundredtrialsb.setStyleSheet(buttonstyleon)
        if self.textbox.isVisible():
            self.textbox.hide()
        if self.label5.isVisible():
            self.label5.setText("")
        if self.label4.isVisible():
            self.label4.hide()
        self.csvbutton = QPushButton("Save to Excel")
        self.csvbutton.setFixedSize(127,57)
        self.csvbutton.setStyleSheet(buttonstyleoff)
        self.page5layout.addWidget(self.csvbutton,3,1)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page5layout.addWidget(self.progress,3,1)
        self.csvbutton.hide()
        self.c = []
        self.label2.setText("")
        self.timer.timeout.connect(lambda: self.update_meter_100_trials())
    def update_meter_100_trials(self):
        self.stop()
        if len(self.c) < 100:
            self.progress.show()
            self.progress.setValue((len(self.c)+1))
            if serial_thread.serial_port.in_waiting:
                serial_data = serial_thread.serial_port.readline().decode().strip()
                if serial_data == '':
                    serial_thread.wait()
                elif serial_data == "inf":
                    self.c.append(float('inf'))
                else:
                    self.c.append(float(serial_data))
        else:
            self.progress.destroy()
            self.progress.deleteLater()
            f = sum(self.c)
            f = str(f)
            if f == 'inf':
                self.label3.setText("Std Deviation: {}\nVariance: {}\nMean: {}".format('inf', 'inf', 'inf'))
                self.timer.timeout.disconnect()
                self.stop()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
            else:
                self.label3.setText("Std Deviation: {:.2f}\nVariance: {:.2f}\nMean: {:.2f}".format(stdev(self.c), variance(self.c), mean(self.c)))
                self.stop()
                self.timer.timeout.disconnect()
                self.csvbutton.show()
                self.textbox.show()
                self.label4.show()
                self.dontsave.show()
                self.dontsave.clicked.connect(lambda: self.savereset())
                self.csvbutton.clicked.connect(lambda: [self.namesave(self.c), self.csvbutton.hide(),self.startbutton.setEnabled(True),self.startbutton.setEnabled(True),self.dontsave.hide(),self.textbox.hide(),self.label4.hide(),self.tentrialsb.setEnabled(True),self.fiftytrialsb.setEnabled(True),self.hundredtrialsb.setEnabled(True)])
                self.tentrialsb.destroy()

    def hideEvent(self, event):
        self.stop()
class Page6(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.label = QLabel("Calibration")
        self.label.setAlignment(Qt.AlignmentFlag.AlignTop)
        fontheader = QFont("Arial", 40)
        font = QFont("Arial", 20)
        self.label.setFont(fontheader)
        self.page6layout = QGridLayout()
        self.blanklabel = QLabel("")
        self.inst = QLabel("1. Connect leads to ohmmeter connections.\n2. Clip leads to each other.\n3. Click the enter button.")
        self.inst.setFont(font)
        self.inst.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.enterbutton = QPushButton("Enter")
        self.enterbutton.setFixedSize(127,57)
        self.enterbutton.setStyleSheet(buttonstyleoff)
        self.page6layout.addWidget(self.blanklabel,0,1)
        self.page6layout.addWidget(self.blanklabel,0,0)
        self.page6layout.addWidget(self.blanklabel,1,0)
        self.page6layout.addWidget(self.enterbutton,1,1)
        self.page6layout.addWidget(self.inst,1,0)

        self.page6layout.addWidget(self.label,0,0)
        self.enterbutton.clicked.connect(self.calibration)
        self.setLayout(self.page6layout)

    def calibration(self):
        self.inst.setText("")
        serial_thread.clear_buffer()
        serial_thread.write_data(b'5')
        serial_thread.start()
        serial_data = serial_thread.serial_port.readline().decode().strip()

        self.enterbutton.hide()
        self.progress = QProgressBar()
        self.progress.setRange(0,100)
        self.progress.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        self.progress.setTextVisible(True)
        self.progress.setValue(0)
        self.progress.setStyleSheet(progessstyle)
        self.page6layout.addWidget(self.progress,1,0)
        self.currentper()
        if serial_data == '1':
            self.completed()
        else:
            self.failed()
        
    def currentper(self):
        a = random.randint(0, 9)
        b = random.randint(10, 19)
        c = random.randint(20, 29)
        d = random.randint(30, 39)
        e = random.randint(40, 49)
        f = random.randint(50, 59)
        g = random.randint(60, 69)
        h = random.randint(70, 79)
        i = random.randint(80, 89)
        j = random.randint(90, 99)
        nums = [a,b,c,d,e,f,g,h,i,j,100]
        for num in nums:
            self.progress.setValue(num)
            QApplication.processEvents()
            time.sleep(.3)
            self.completed()
    def completed(self):
        self.progress.destroy()
        self.progress.deleteLater()
        self.inst.setText("Calibration Successful")
    def failed(self):
        self.progress.destroy()
        self.progress.deleteLater()
        self.inst.setText("Calibration Unsuccessful")
    def hideEvent(self, event):
        self.reset()
    def reset(self):
        self.inst.setText("1. Connect leads to ohmmeter connections.\n2. Clip leads to each other.\n3. Click the enter button.")
        self.enterbutton.show()
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # arduino.connect('COM14')
        self.setWindowTitle("E-Textile UI")
        self.setGeometry(100, 100, 800, 480)
        self.button1 = QPushButton("Home")
        self.button2 = QPushButton("Voltmeter")
        self.button3 = QPushButton("Diode/LED Test")
        self.button4 = QPushButton("Ohmmeter")
        self.button5 = QPushButton("Continuity Test")
        self.button6 = QPushButton("Calibration")

        self.button1.clicked.connect(lambda: [self.showPage1(),serial_thread.terminate()])
        self.button2.clicked.connect(lambda: [self.showPage2(),serial_thread.terminate()])
        self.button3.clicked.connect(lambda: [self.showPage3(),serial_thread.terminate()])
        self.button4.clicked.connect(lambda: [self.showPage4(),serial_thread.terminate()])
        self.button5.clicked.connect(lambda: [self.showPage5(),serial_thread.terminate()])
        self.button6.clicked.connect(lambda: [self.showPage6(),serial_thread.terminate()])
        buttonLayout = QHBoxLayout()
        buttonLayout.addWidget(self.button1)
        buttonLayout.addWidget(self.button2)
        buttonLayout.addWidget(self.button3)
        buttonLayout.addWidget(self.button4)
        buttonLayout.addWidget(self.button5)
        buttonLayout.addWidget(self.button6)
        mainLayout = QVBoxLayout()
        mainLayout.addLayout(buttonLayout)
        self.stackedLayout = QStackedLayout()
        self.stackedLayout.addWidget(Page1())
        self.stackedLayout.addWidget(Page2())
        self.stackedLayout.addWidget(Page3())
        self.stackedLayout.addWidget(Page4())
        self.stackedLayout.addWidget(Page5())
        self.stackedLayout.addWidget(Page6())
        mainLayout.addLayout(self.stackedLayout)
        self.setLayout(mainLayout)

        # Set button style
        self.button1.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.button2.setStyleSheet(buttonstyleoff)
        self.button3.setStyleSheet(buttonstyleoff)
        self.button4.setStyleSheet(buttonstyleoff)
        self.button5.setStyleSheet(buttonstyleoff)
        self.button6.setStyleSheet(buttonstyleoff)

    def showPage1(self):
        self.stackedLayout.setCurrentIndex(0)
        self.button1.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.button2.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button3.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button4.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button5.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button6.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")

    def showPage2(self):
        self.stackedLayout.setCurrentIndex(1)
        self.button1.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button2.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.button3.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button4.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button5.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button6.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")

    def showPage3(self):
        self.stackedLayout.setCurrentIndex(2)
        self.button1.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button2.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button3.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.button4.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button5.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button6.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")

    def showPage4(self):
        self.stackedLayout.setCurrentIndex(3)
        self.button1.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button2.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button3.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button4.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.button5.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button6.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")

    def showPage5(self):
        self.stackedLayout.setCurrentIndex(4)
        self.button1.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button2.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button3.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button4.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button5.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")
        self.button6.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")

    def showPage6(self):
        self.stackedLayout.setCurrentIndex(5)
        self.button1.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button2.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button3.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button4.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button5.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #DFE3E3;")
        self.button6.setStyleSheet("border-radius: 10px;padding: 20px;background-color: #D7AB29;")

    # ser = serial.Serial('COM14', 9600)
    # while True:
    #     if ser.in_waiting > 0:
    #         data = ser.readline().decode().rstrip()
    #         print(data)

if __name__ == "__main__":
    arduinoport = detect_arduino_port()
    serial_thread = SerialThread(arduinoport)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
