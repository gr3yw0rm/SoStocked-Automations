import PyQt5.QtWidgets as qtw
import PyQt5.QtGui as qtg

# https://www.youtube.com/watch?v=O58FGYYBV7U&list=PLCC34OHNcOtpmCA8s_dpPMvQLyHbvxocY&index=2&ab_channel=Codemy.com
# Open PyQt Designer GUI by command "designer.exe"
# convert the .ui file to py file "pyuic5 -x hello_world.ui -o hello_world.py"

class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()
        # Add a title
        self.setWindowTitle("SoStocked Automation")

        # Set a layout
        self.setLayout(qtw.QVBoxLayout())

        # Create a label
        my_label = qtw.QLabel("Pick Something from the list")
        # Change the font size of label
        my_label.setFont(qtg.QFont('Helvetica', 24))
        self.layout().addWidget(my_label)
        
        # Create an entry box
        my_entry = qtw.QLineEdit()
        my_entry.setObjectName("name_field")
        my_entry.setText("")
        self.layout().addWidget(my_entry)

        # Create an combo box (drop down)
        my_combo = qtw.QComboBox(self)
        # Add items to the combo box
        my_combo.addItem("Peperroni")
        my_combo.addItem("Cheese")
        my_combo.addItem("Mushroom")
        my_combo.addItem("Peppers")



        # Create a button
        my_button = qtw.QPushButton("Press Me!", 
                        clicked = lambda: press_it())
        self.layout().addWidget(my_button)

        self.show()

        def press_it():
            # Add name to label
            my_label.setText(f'Hello {my_entry.text()}!')
            # Clear the entry box 
            my_entry.setText("")        


app = qtw.QApplication([])
mw = MainWindow()

# Run the App
app.exec_()