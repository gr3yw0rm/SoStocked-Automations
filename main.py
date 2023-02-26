from ast import arg
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QCommandLinkButton, QFileDialog
from PyQt5 import uic
from PyQt5 import QtGui
import os
import sys
from sostocked_templates import update_inventory, send_to_amazon
from amazon_packinglist import create_shippinguploads



class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()

        # Load the ui file
        uic.loadUi("sostocked_automation.ui", self)
        self.setWindowIcon(QtGui.QIcon('sostocked.ico'))

        # Define our widgets
        self.button   = self.findChild(QPushButton, "pushButton")
        self.button_2 = self.findChild(QPushButton, "pushButton_2")
        self.button_3 = self.findChild(QPushButton, "pushButton_3")
        self.label    = self.findChild(QLabel, "label") # tab 1 step 1
        self.label_2  = self.findChild(QLabel, "label_2") # open file
        self.label_3  = self.findChild(QLabel, "label_3") # saved to
        self.label_4  = self.findChild(QLabel, "label_4") # step 3
        self.label_5  = self.findChild(QLabel, "label_5") # tab 2 - step 1
        self.label_6  = self.findChild(QLabel, "label_6") # open file
        self.label_7  = self.findChild(QLabel, "label_7") # saved to
        self.label_8  = self.findChild(QLabel, "label_8") # step 3
        self.label_9  = self.findChild(QLabel, "label_9") # step 4
        self.label_10  = self.findChild(QLabel, "label_10") # open file 2
        self.label_11  = self.findChild(QLabel, "label_11") # saved to
        self.label_12  = self.findChild(QLabel, "label_12") # step 5
        self.convert   = self.findChild(QCommandLinkButton, "commandLinkButton")
        self.convert_2 = self.findChild(QCommandLinkButton, "commandLinkButton_2")
        self.convert_3 = self.findChild(QCommandLinkButton, "commandLinkButton_3")

        # Click the Dropdown Box
        self.button.clicked.connect(self.select_shopify_inventory)
        self.button_2.clicked.connect(self.select_sostocked_shipment)
        self.button_3.clicked.connect(self.select_amazon_packlist)

        # Convert excel files
        self.convert.clicked.connect(self.convert_shopify)
        self.convert_2.clicked.connect(self.convert_sostocked)
        self.convert_3.clicked.connect(self.convert_shipmentPacklist)

        # Downloads Folder
        self.downloadsDirectory = "C:\\" + os.path.join(os.getenv('HOMEPATH'), 'Downloads')

        # Show the App
        self.show()


    # Open File Dialogs
    def select_shopify_inventory(self):
        fname = QFileDialog.getOpenFileName(self, "Open File", self.downloadsDirectory, "All Files (*)")
        if fname:
            # Output filename to screen
            self.label_2.setText(fname[0])
            self.shopify_inventory = fname[0]

    def select_sostocked_shipment(self):
        fname = QFileDialog.getOpenFileName(self, "Open File", self.downloadsDirectory, "All Files (*)")
        if fname:
            self.label_6.setText(fname[0])
            self.sostocked_shipment = fname[0]

    def select_amazon_packlist(self):
        fname = QFileDialog.getOpenFileName(self, "Open File", self.downloadsDirectory, "All Files (*)")
        if fname:
            self.label_10.setText(fname[0])
            self.amazon_packlist = fname[0]

    # Conversion of files
    def convert_shopify(self):
        """Converts Shipping Tree from Amazon Analytics Report (Units Sold) to SoStocked's bulk import template"""
        try:
            saved_location = update_inventory(self.shopify_inventory)
        except Exception as error:
            saved_location = f'ERROR: {error}'
        self.label_3.setText(f'Saved to: /{saved_location}')

    def convert_sostocked(self):
        """Converts SoStocked's transfer forecast excel to Amazon manifest file template upload"""
        try:
            saved_location = send_to_amazon(self.sostocked_shipment)
        except Exception as error:
            saved_location = f'ERROR: {error}'
        self.label_7.setText(f'Saved to: /{saved_location}')

    def convert_shipmentPacklist(self):
        """Scrapes & converts Amazon shipment packing list to SoStocked import shipment & ST upload files"""
        try:
            sostocked_saved_location = create_shippinguploads(self.amazon_packlist)
            st_saved_location = ""
            saved_location = f"{sostocked_saved_location}>>{st_saved_location}"
        except Exception as error:
            saved_location = f'ERROR: {error}'
        self.label_11.setText(f'Saved to: /{saved_location}')


# Initialize the App
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('sostocked.ico'))
    # app_icon = Q
    UIWindow = UI()
    app.exec()
    print(__name__)
    # run pyqt5 designer
    # $ designer
    # $ pyinstaller 'sostocked automations.spec'
    # $ pyinstaller --onefile --noconsole --ico=sostocked.ico --name "sostocked automations" main.py