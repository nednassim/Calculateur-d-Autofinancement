import sys
from datetime import datetime
import pandas as pd
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QScrollArea, QHeaderView,
    QSizePolicy, QApplication
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QTextDocument
from PySide6.QtPrintSupport import QPrinter
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class AutofinancementCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Calculateur d'Autofinancement")
        self.setMinimumSize(800, 600)
        
        # Setup main scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.setCentralWidget(scroll)
        
        main_widget = QWidget()
        scroll.setWidget(main_widget)
        
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)
        
        # Setup UI components
        self.setup_ui(main_layout)
    
    def setup_ui(self, main_layout):
        # Title
        title_label = QLabel("Calculateur d'Autofinancement")
        title_label.setObjectName("title")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        main_layout.addWidget(title_label)
        
        # Input section
        input_frame = self.create_input_frame()
        main_layout.addWidget(input_frame)
        
        # Calculate button
        self.calculate_btn = QPushButton("Calculer")
        self.calculate_btn.setObjectName("calculateButton")
        self.calculate_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        main_layout.addWidget(self.calculate_btn, alignment=Qt.AlignCenter)
        
        # Results section
        results_frame = self.create_results_frame()
        main_layout.addWidget(results_frame)
        
        # Connect signals
        self.calculate_btn.clicked.connect(self.calculate)
    
    def create_input_frame(self):
        frame = QFrame()
        frame.setObjectName("inputFrame")
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        
        # Import button
        self.import_excel = QPushButton("Importer Excel")
        self.import_excel.setObjectName("importButton")
        self.import_excel.clicked.connect(self.import_xlsx_data)
        layout.addWidget(self.import_excel)
        
        # Table with scroll
        table_scroll = QScrollArea()
        table_scroll.setWidgetResizable(True)
        
        self.input_table = QTableWidget()
        self.setup_table()
        
        table_scroll.setWidget(self.input_table)
        layout.addWidget(table_scroll)
        
        return frame
    
    def setup_table(self):
        self.input_table.setColumnCount(2)
        self.input_table.setHorizontalHeaderLabels(["Élément", "Montant (DZD)"])
        self.input_table.setRowCount(20)
        
        elements = [
            "Ventes et produits annexes",
            "Variation stocks produits finis et en cours",
            "Production immobilisée",
            "Subventions d'exploitation",
            "Achats consommés",
            "Services extérieurs et autres consommations",
            "Charges de personnel",
            "Impôts, taxes et versements assimilés",
            "Autres produits opérationnels",
            "Autres charges opérationnelles",
            "Dotations aux amortissements, provisions et pertes de valcur",
            "Reprise sur pertes de valcur et provisions",
            "Produits financiers",
            "Charges financières",
            "Impôts exigibles sur résultats ordinaires",
            "Impôts différés (Variations) sur résultats ordinaires",
            "Eléments extraordinaires (produits)",
            "Eléments extraordinaires (charges)",
            "Part dans les résultats nets des sociétés mises en équivalence",
            "Dividendes versés"
        ]
        
        for i, element in enumerate(elements):
            self.input_table.setItem(i, 0, QTableWidgetItem(element))
            self.input_table.setItem(i, 1, QTableWidgetItem("0"))
            self.input_table.item(i, 0).setFlags(Qt.ItemIsEnabled)
        
        # Table sizing
        self.input_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.input_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.input_table.verticalHeader().setDefaultSectionSize(30)
        self.input_table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
    
    def create_results_frame(self):
        frame = QFrame()
        frame.setObjectName("resultsFrame")
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(20)
        
        # Key results with scroll
        key_results_scroll = QScrollArea()
        key_results_scroll.setWidgetResizable(True)
        
        key_results = QWidget()
        key_results_layout = QVBoxLayout(key_results)
        key_results_layout.setContentsMargins(5, 5, 5, 5)
        key_results_layout.setSpacing(15)
        
        # Result labels
        self.resultat_net_label = QLabel("Résultat net de l'exercice")
        self.resultat_net_label.setObjectName("resultLabel")
        self.resultat_net_value = QLabel("0 DZD")
        self.resultat_net_value.setObjectName("resultValue")
        
        self.caf_label = QLabel("Capacité d'Autofinancement (CAF)")
        self.caf_label.setObjectName("resultLabel")
        self.caf_value = QLabel("0 DZD")
        self.caf_value.setObjectName("resultValue")
        
        self.autofinancement_label = QLabel("Autofinancement")
        self.autofinancement_label.setObjectName("resultLabel")
        self.autofinancement_value = QLabel("0 DZD")
        self.autofinancement_value.setObjectName("resultValue")
        
        self.interpretation_label = QLabel("Interprétation")
        self.interpretation_label.setObjectName("resultLabel")
        self.interpretation_value = QLabel("")
        self.interpretation_value.setObjectName("interpretationText")
        self.interpretation_value.setWordWrap(True)
        
        key_results_layout.addWidget(self.resultat_net_label)
        key_results_layout.addWidget(self.resultat_net_value)
        key_results_layout.addWidget(self.caf_label)
        key_results_layout.addWidget(self.caf_value)
        key_results_layout.addWidget(self.autofinancement_label)
        key_results_layout.addWidget(self.autofinancement_value)
        key_results_layout.addWidget(self.interpretation_label)
        key_results_layout.addWidget(self.interpretation_value)
        
        key_results_scroll.setWidget(key_results)
        layout.addWidget(key_results_scroll, stretch=1)
        
        # Chart
        chart_box = self.create_chart_box()
        layout.addWidget(chart_box, stretch=1)
        
        return frame
    
    def create_chart_box(self):
        box = QFrame()
        box.setObjectName("chartBox")
        box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        layout = QVBoxLayout(box)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        
        chart_title = QLabel("Visualisation")
        chart_title.setObjectName("chartTitle")
        layout.addWidget(chart_title, alignment=Qt.AlignCenter)
        
        self.figure = Figure(figsize=(5, 4), dpi=100, tight_layout=True)
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.canvas)
        
        self.export_btn = QPushButton("Exporter PDF")
        self.export_btn.setObjectName("exportButton")
        self.export_btn.setFixedSize(150, 40)
        self.export_btn.clicked.connect(self.export_to_pdf)
        layout.addWidget(self.export_btn, alignment=Qt.AlignCenter)
        
        return box
    
    def import_xlsx_data(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "Importer un fichier Excel",
            "",
            "Fichiers Excel (*.xlsx *.xls)"
        )
        
        if not filepath:
            return
        
        try:
            df = pd.read_excel(filepath)
            
            # Normalize column names
            df.columns = [col.strip().lower().replace('é', 'e').replace('è', 'e') for col in df.columns]
            
            if not all(col in df.columns for col in ['libelle', 'montant']):
                raise ValueError("Colonnes requises: 'Libellé' et 'Montant'")
            
            data = {}
            for _, row in df.iterrows():
                element = str(row['libelle']).strip()
                if not element or element in data:
                    continue
                try:
                    value = float(row['montant'])
                    data[element] = value
                except (ValueError, KeyError):
                    continue
            
            # Update table while maintaining size
            for row in range(self.input_table.rowCount()):
                element = self.input_table.item(row, 0).text()
                matching_key = None
                for key in data.keys():
                    if element.lower().strip() == key.lower().strip():
                        matching_key = key
                        break
                
                if matching_key:
                    self.input_table.setItem(row, 1, QTableWidgetItem(str(data[matching_key])))
            
            QMessageBox.information(self, "Succès", "Données importées avec succès!")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur d'importation:\n{str(e)}")
    
    def calculate(self):
        try:
            values = {}
            for row in range(self.input_table.rowCount()):
                element = self.input_table.item(row, 0).text()
                value = self.input_table.item(row, 1).text()
                values[element] = float(value) if value else 0.0
            
            resultat_net = values["Part dans les résultats nets des sociétés mises en équivalence"]

            # Calculate results
            caf = (
                resultat_net +
                values["Ventes et produits annexes"] -
                values["Variation stocks produits finis et en cours"] -
                values["Production immobilisée"] -
                values["Subventions d'exploitation"] +
                values["Achats consommés"] +
                values["Services extérieurs et autres consommations"] +
                values["Charges de personnel"] +
                values["Impôts, taxes et versements assimilés"] -
                values["Autres produits opérationnels"] +
                values["Autres charges opérationnelles"] +
                values["Dotations aux amortissements, provisions et pertes de valcur"] -
                values["Reprise sur pertes de valcur et provisions"] -
                values["Produits financiers"] +
                values["Charges financières"] +
                values["Impôts exigibles sur résultats ordinaires"] -
                values["Impôts différés (Variations) sur résultats ordinaires"] -
                values["Eléments extraordinaires (produits)"] +
                values["Eléments extraordinaires (charges)"] 
            )
            
            autofinancement = caf - values["Dividendes versés"]
            
            # Update UI
            self.resultat_net_value.setText(f"{resultat_net:,.2f} DZD")
            self.caf_value.setText(f"{caf:,.2f} DZD")
            self.autofinancement_value.setText(f"{autofinancement:,.2f} DZD")
            
            self.interpretation_value.setText(self.get_interpretation(resultat_net, caf, autofinancement))
            self.update_chart(resultat_net, caf, autofinancement)
            
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur de calcul:\n{str(e)}")
    
    def get_interpretation(self, resultat_net, caf, autofinancement):
        interpretations = []
        
        if resultat_net > 0:
            interpretations.append(f"Résultat net positif: {resultat_net:,.2f} DZD")
        else:
            interpretations.append(f"Résultat net négatif: {resultat_net:,.2f} DZD")
        
        if caf > 0:
            interpretations.append(f"CAF positive: {caf:,.2f} DZD")
        else:
            interpretations.append(f"CAF négative: {caf:,.2f} DZD")
        
        if autofinancement > 0:
            interpretations.append(f"Autofinancement positif: {autofinancement:,.2f} DZD")
        else:
            interpretations.append(f"Autofinancement négatif: {autofinancement:,.2f} DZD")
        
        return "\n\n".join(interpretations)
    
    def update_chart(self, resultat_net, caf, autofinancement):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        labels = ['Résultat Net', 'CAF', 'Autofinancement']
        values = [resultat_net, caf, autofinancement]
        colors = ['#FF5722', '#2196F3', '#4CAF50']
        
        width = 0.6
        x_pos = range(len(values))
        bars = ax.bar(x_pos, values, width, color=colors)
        
        ax.set_xticks(x_pos)
        ax.set_xticklabels(labels, rotation=15, ha='right')
        ax.yaxis.set_major_formatter('DZD{x:,.0f}')
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                    f'DZD{height:,.0f}',
                    ha='center', va='bottom', fontsize=10)
        
        y_min = min(0, min(values)*1.1)
        y_max = max(values)*1.2
        ax.set_ylim(y_min, y_max)
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(axis='y', alpha=0.3)
        
        self.figure.tight_layout()
        self.canvas.draw()
    
    def export_to_pdf(self):
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Exporter en PDF", "", "PDF Files (*.pdf)"
        )
        if filepath:
            try:
                chart_path = "temp_chart.png"
                self.figure.savefig(chart_path, bbox_inches='tight', dpi=150)
                
                printer = QPrinter(QPrinter.HighResolution)
                printer.setOutputFormat(QPrinter.PdfFormat)
                printer.setOutputFileName(filepath)
                
                doc = QTextDocument()
                html = self.generate_report_html(chart_path)
                doc.setHtml(html)
                
                doc.print_(printer)
                
                QMessageBox.information(self, "Succès", "Rapport exporté avec succès!")
                
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Échec de l'export:\n{str(e)}")
    
    def generate_report_html(self, chart_path):
        return f"""
        <html>
        <head>
        <style>
        body {{ font-family: Arial; margin: 20px; }}
        h1 {{ color: #333; border-bottom: 1px solid #eee; padding-bottom: 10px; }}
        .header {{ background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
        .result {{ margin: 15px 0; }}
        .value {{ font-weight: bold; color: #2196F3; }}
        .interpretation {{ background-color: #f9f9f9; padding: 15px; border-left: 4px solid #2196F3; border-radius: 4px; }}
        .chart-container {{ text-align: center; margin: 20px 0; }}
        </style>
        </head>
        <body>
        <div class="header">
            <h1>Rapport d'Autofinancement</h1>
            <p>Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
        </div>
        
        <div class="results">
            <h2>Résultats Clés</h2>
            <div class="result">
                <span>Résultat net de l'exercice: </span>
                <span class="value">{self.resultat_net_value.text()}</span>
            </div>
            <div class="result">
                <span>Capacité d'Autofinancement (CAF): </span>
                <span class="value">{self.caf_value.text()}</span>
            </div>
            <div class="result">
                <span>Autofinancement: </span>
                <span class="value">{self.autofinancement_value.text()}</span>
            </div>
        </div>
        
        <div class="chart-container">
            <h2>Visualisation</h2>
            <img src="{chart_path}" width="500" />
        </div>
        
        <div class="interpretation">
            <h2>Interprétation</h2>
            <p>{self.interpretation_value.text().replace('\n\n', '<br><br>')}</p>
        </div>
        </body>
        </html>
        """