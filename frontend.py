import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QInputDialog
import backend  # Assuming backend handles data operations

class ExpenseTracker(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Expense Tracker')
        self.setGeometry(300, 300, 600, 400)

        # Main Menu Buttons
        self.add_button = QPushButton('Add Expense', self)
        self.add_button.clicked.connect(self.add_expense)

        self.view_button = QPushButton('View Expenses', self)
        self.view_button.clicked.connect(self.view_expenses)

        self.save_button = QPushButton('Save to Excel', self)
        self.save_button.clicked.connect(self.save_expenses)
        self.save_button.move(50, 50)

        self.load_button = QPushButton('Load from Excel', self)
        self.load_button.clicked.connect(self.load_expenses)

        self.clear_button = QPushButton('Clear Data', self)
        self.clear_button.clicked.connect(self.clear_data)
        self.clear_button.move(50, 100)

        self.exit_button = QPushButton('Exit', self)
        self.exit_button.clicked.connect(self.close)

        self.show()

        # Layout Setup
        layout = QVBoxLayout()
        layout.addWidget(self.add_button)
        layout.addWidget(self.view_button)
        layout.addWidget(self.save_button)
        layout.addWidget(self.load_button)
        layout.addWidget(self.clear_button)
        layout.addWidget(self.exit_button)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        central_widget.setLayout(layout)

    def add_expense(self):
        date, date_ok = QInputDialog.getText(self, "Input Date", "Enter the date (DD-MM-YYYY):")
        if date_ok:
            category, category_ok = QInputDialog.getText(self, "Input Category", "Enter the category:")
            if category_ok:
                amount, amount_ok = QInputDialog.getDouble(self, "Input Amount", "Enter the amount:")
                if amount_ok:
                    backend.add_expense(date, category, amount)
                    QMessageBox.information(self, 'Success', 'Expense added successfully.')

    def view_expenses(self):
        choice, ok = QInputDialog.getItem(self, "View Expenses", "Choose the type of view:",
                                          ["Monthly Expenses", "Total Expenses"])
        if ok and choice == "Monthly Expenses":
            self.view_monthly_expenses()
        elif ok and choice == "Total Expenses":
            self.view_total_expenses()

    def view_monthly_expenses(self):
        month, ok = QInputDialog.getText(self, "Monthly Expenses", "Enter the month (MM-YYYY):")
        if ok:
            report = backend.view_monthly_expenses(month)
            QMessageBox.information(self, f'Expenses for {month}', report)

    def view_total_expenses(self):
        report = backend.view_total_expenses()
        QMessageBox.information(self, 'Annual Summary', report)

    def save_expenses(self):
        filepath, _ = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel Files (*.xlsx);;All Files (*)')
        if filepath:
            message = backend.save_to_excel(filepath)
            QMessageBox.information(self, 'Save to Excel', message)
        else:
            QMessageBox.warning(self, 'Save to Excel', 'No file was selected.')

    def load_expenses(self):
        # Open a file dialog to select the Excel file
        filepath, _ = QFileDialog.getOpenFileName(self, 'Open File', '', 'Excel Files (*.xlsx);;All Files (*)')
        if filepath:
            # Call backend function to load data from the selected Excel file
            result = backend.load_from_excel(filepath)
            QMessageBox.information(self, 'Load Data', result)
        else:
            QMessageBox.warning(self, 'Load Data', 'No file was selected.')

    def clear_data(self):
        option, ok = QInputDialog.getItem(self, "Clear Data", "Choose data to clear:", ["All", "Month", "Date"])
        if ok:
            if option == "All":
                backend.clear_data('all')
                QMessageBox.information(self, 'Clear Data', 'All data has been cleared.')
            elif option == "Month":
                month_year, ok = QInputDialog.getText(self, "Enter Month and Year", "Enter month and year (MM-YYYY):")
                if ok:
                    backend.clear_data('month', month_year)
                    QMessageBox.information(self, 'Clear Data', f'All data for {month_year} has been cleared.')
            elif option == "Date":
                date, ok = QInputDialog.getText(self, "Enter Date", "Enter date (DD-MM-YYYY):")
                if ok:
                    backend.clear_data('date', date)
                    QMessageBox.information(self, 'Clear Data', f'All data for {date} has been cleared.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExpenseTracker()
    sys.exit(app.exec_())
