import sys
from PySide6.QtWidgets import QApplication
from calculator import AutofinancementCalculator

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Load the stylesheet
    with open("style.css", "r") as f:
        app.setStyleSheet(f.read())
    
    calculator = AutofinancementCalculator()
    calculator.show()
    sys.exit(app.exec())