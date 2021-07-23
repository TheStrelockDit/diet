import sys

from controller import *
from interface import *

app = QApplication(sys.argv)
ex = MainMenu()
sys.exit(app.exec_())


