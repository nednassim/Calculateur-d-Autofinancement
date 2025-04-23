import sys
from PySide6.QtWidgets import QApplication
from calculator import AutofinancementCalculator

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Load stylesheet
    try:
        with open("style.css", "r") as f:
            app.setStyleSheet(f.read())
    except Exception as e:
        print(f"Could not load stylesheet: {e}")
    
    calculator = AutofinancementCalculator()
    calculator.show()
    sys.exit(app.exec())