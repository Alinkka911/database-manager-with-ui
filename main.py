from PyQt5.QtWidgets import QApplication
from ui import LoginWindow
from auth import add_user

if __name__ == "__main__":
    add_user("User", "user1", "user")
    app = QApplication([])
    window = LoginWindow()
    window.show()
    app.exec_()