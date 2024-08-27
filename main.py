from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QFileDialog, QPushButton, QLineEdit
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
import pandas as pd

class DataViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Consulta")
        self.setGeometry(100, 100, 1000, 600)
        
        # Configurar el icono
        self.setWindowIcon(QIcon("C:/Users/Ava Montajes/Documents/APPCONSULTA/favicon.ico"))
        
        # Crear widgets
        self.table = QTableWidget()
        self.table.setHorizontalHeaderLabels([])
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar trabajador...")
        
        self.search_button = QPushButton("Buscar")
        self.search_button.clicked.connect(self.search_worker)
        
        self.load_button = QPushButton("Cargar Excel")
        self.load_button.clicked.connect(self.load_file)
        
        # Crear el layout
        layout = QVBoxLayout()
        layout.addWidget(self.load_button)
        layout.addWidget(self.search_input)
        layout.addWidget(self.search_button)
        layout.addWidget(self.table)
        
        # Crear el widget central
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        self.df = pd.DataFrame()
        
        # Aplicar estilos
        self.apply_styles()
        
    def apply_styles(self):
        # Aplicar estilo CSS
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QLineEdit {
                border: 1px solid #ced4da;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #ced4da;
                border-radius: 5px;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #007bff;
                color: white;
                padding: 10px;
            }
        """)
        
    def load_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx; *.xls)", options=options)
        
        if file_path:
            try:
                self.df = pd.read_excel(file_path, header=6)
                self.update_table(self.df)
            except Exception as e:
                print(f"Error loading file: {e}")
        
    def update_table(self, df):
        self.table.setRowCount(0)
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.tolist())
        
        for row_index, row in df.iterrows():
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(str(value) if pd.notna(value) else '')
                item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_position, column_index, item)
                
        self.adjust_column_widths()

    def adjust_column_widths(self):
        for column in range(self.table.columnCount()):
            self.table.resizeColumnToContents(column)
        
    def search_worker(self):
        search_name = self.search_input.text().strip()
        if search_name:
            filtered_df = self.df[self.df['NOMBRE'].str.contains(search_name, case=False, na=False)]
            self.update_table(filtered_df)
        else:
            self.update_table(self.df)

app = QApplication([])
window = DataViewer()
window.show()
app.exec_()
