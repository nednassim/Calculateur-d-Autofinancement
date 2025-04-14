import sys
from datetime import datetime
import pandas as pd
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QPushButton, QComboBox, QLineEdit, QFrame, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QMenuBar, QStatusBar, QApplication
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPixmap, QIcon, QPainter, QTextDocument
from PySide6.QtPrintSupport import QPrinter
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class AutofinancementCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Calculateur d'Autofinancement")
        self.resize(1000, 800)
        self.recent_files = []
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # Setup menus and status bar
        self.setup_menubar()
        self.setup_status_bar()
        
        # Title
        title_label = QLabel("Calculateur d'Autofinancement")
        title_label.setObjectName("title")
        main_layout.addWidget(title_label)
        
        # Data input section
        input_frame = QFrame()
        input_frame.setObjectName("inputFrame")
        input_layout = QVBoxLayout(input_frame)
        
        # Import buttons
        import_buttons = QHBoxLayout()
        self.import_csv = QPushButton("Importer CSV")
        self.import_csv.setObjectName("importButton")
        self.import_csv.clicked.connect(self.import_csv_data)
        
        self.import_excel = QPushButton("Importer Excel")
        self.import_excel.setObjectName("importButton")
        self.import_excel.clicked.connect(self.import_xlsx_data)
        
        import_buttons.addWidget(self.import_csv)
        import_buttons.addWidget(self.import_excel)
        import_buttons.addStretch()
        input_layout.addLayout(import_buttons)
        
        # Input table
        self.input_table = QTableWidget()
        self.input_table.setColumnCount(2)
        self.input_table.setHorizontalHeaderLabels(["Élément", "Montant (€)"])
        self.input_table.setRowCount(10)
        self.input_table.setMinimumHeight(250)  # Hauteur minimale en pixels
        #self.input_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Set up default rows with tooltips
        elements = [
            ("Résultat Net", "Bénéfice net après impôts"),
            ("Dotations aux amortissements", "Charges non décaissables comptabilisées"),
            ("Variation des stocks", "Variation des stocks de matières premières et produits finis"),
            ("Variation des créances clients", "Variation des créances clients et autres créances"),
            ("Variation des dettes fournisseurs", "Variation des dettes fournisseurs et autres dettes"),
            ("Autres produits encaissables", "Autres produits qui se traduiront par des encaissements"),
            ("Autres charges décaissables", "Autres charges qui se traduiront par des décaissements"),
            ("Dividendes versés", "Dividendes distribués aux actionnaires"),
            ("Investissements", "Total des investissements réalisés"),
            ("Désinvestissements", "Produits des cessions d'actifs")
        ]
        
        for i, (element, tooltip) in enumerate(elements):
            self.input_table.setItem(i, 0, QTableWidgetItem(element))
            self.input_table.setItem(i, 1, QTableWidgetItem("0"))
            self.input_table.item(i, 0).setFlags(Qt.ItemIsEnabled)
            self.input_table.item(i, 0).setToolTip(tooltip)
            self.input_table.item(i, 1).setToolTip(tooltip)
        
        self.input_table.horizontalHeader().setStretchLastSection(True)
        input_layout.addWidget(self.input_table)
        
        main_layout.addWidget(input_frame)
        
        # Calculate button
        self.calculate_btn = QPushButton("Calculer")
        self.calculate_btn.setObjectName("calculateButton")
        main_layout.addWidget(self.calculate_btn, alignment=Qt.AlignCenter)
        
        # Results section
        results_frame = QFrame()
        results_frame.setObjectName("resultsFrame")
        results_layout = QHBoxLayout(results_frame)
        
        # Key results
        key_results = QVBoxLayout()
        key_results.setSpacing(15)
        
        self.caf_label = QLabel("Capacité d'Autofinancement (CAF)")
        self.caf_label.setObjectName("resultLabel")
        self.caf_value = QLabel("0 €")
        self.caf_value.setObjectName("resultValue")
        
        self.autofinancement_label = QLabel("Autofinancement")
        self.autofinancement_label.setObjectName("resultLabel")
        self.autofinancement_value = QLabel("0 €")
        self.autofinancement_value.setObjectName("resultValue")
        
        self.taux_label = QLabel("Taux d'Autofinancement")
        self.taux_label.setObjectName("resultLabel")
        self.taux_value = QLabel("0 %")
        self.taux_value.setObjectName("resultValue")
        
        self.interpretation_label = QLabel("Interprétation")
        self.interpretation_label.setObjectName("resultLabel")
        self.interpretation_value = QLabel("")
        self.interpretation_value.setObjectName("interpretationText")
        self.interpretation_value.setWordWrap(True)
        
        key_results.addWidget(self.caf_label)
        key_results.addWidget(self.caf_value)
        key_results.addWidget(self.autofinancement_label)
        key_results.addWidget(self.autofinancement_value)
        key_results.addWidget(self.taux_label)
        key_results.addWidget(self.taux_value)
        key_results.addWidget(self.interpretation_label)
        key_results.addWidget(self.interpretation_value)
        
        results_layout.addLayout(key_results)
        
        # Chart visualization
        self.figure = Figure(figsize=(3, 4), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        
        chart_box = QFrame()
        chart_box.setObjectName("chartBox")
        chart_layout = QVBoxLayout(chart_box)
        
        chart_title = QLabel("Visualisation des Résultats")
        chart_title.setObjectName("chartTitle")
        chart_layout.addWidget(chart_title)
        chart_layout.addWidget(self.canvas)
        
        export_container = QFrame()
        export_layout = QHBoxLayout(export_container)
        export_layout.setContentsMargins(0, 10, 0, 0)  # Espacement supérieur

        self.export_btn = QPushButton("Exporter en PDF")
        self.export_btn.setObjectName("exportButton")
        #self.export_btn.setIcon(QIcon("assets/export_icon.png"))
        self.export_btn.setFixedSize(180, 60)
        export_layout.addWidget(self.export_btn, alignment=Qt.AlignCenter)


        chart_layout.addWidget(export_container)
        
        results_layout.addWidget(chart_box)
        main_layout.addWidget(results_frame)
        
        # Connect signals
        self.setup_connections()
        
        # Load sample data for demonstration
        self.load_sample_data()
        
        # Initial calculation
        self.calculate()
    
    def import_csv_data(self):
        """Importe un fichier CSV avec le format exact spécifié"""
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "Importer un fichier financier",
            "",
            "Fichiers CSV (*.csv)"
        )
        
        if not filepath:
            return
        
        try:
            # Lecture du CSV avec encoding UTF-8 et vérification du format
            df = pd.read_csv(filepath, encoding='utf-8')
            
            # Vérification des colonnes obligatoires
            if not all(col in df.columns for col in ['Élément', 'Montant']):
                raise ValueError("Format de fichier invalide. Colonnes requises: 'Élément', 'Montant'")
            
            # Conversion en dictionnaire {Élément: Montant}
            data = dict(zip(df['Élément'], df['Montant']))
            
            # Mapping exact vers les champs du tableau
            field_mapping = [
                "Résultat Net",
                "Dotations aux amortissements",
                "Variation des stocks",
                "Variation des créances clients",
                "Variation des dettes fournisseurs",
                "Autres produits encaissables",
                "Autres charges décaissables",
                "Dividendes versés",
                "Investissements",
                "Désinvestissements"
            ]
            
            # Mise à jour de l'interface
            for row in range(self.input_table.rowCount()):
                field_name = self.input_table.item(row, 0).text()
                if field_name in data:
                    value = str(data[field_name])
                    item = QTableWidgetItem(value)
                    item.setTextAlignment(Qt.AlignCenter)

                    self.input_table.setItem(row, 1, item)
            
            # Rafraîchissement de l'UI
            self.input_table.viewport().update()
            QMessageBox.information(self, "Succès", "Données importées avec succès!")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'importation", 
                f"Le fichier doit respecter exactement ce format:\n\n"
                f"Élément,Montant\n"
                f"Résultat Net,150000\n"
                f"Dotations aux amortissements,50000\n"
                f"...\n\n"
                f"Erreur technique: {str(e)}"
            )


    def import_xlsx_data(self):
        filepath, _ = QFileDialog.getOpenFileName(
        self,
        "Import Excel File",
        "",
        "Excel Files (*.xlsx *.xls)"
         )
        
        if not filepath:
            return
        
        try:
            # Read Excel (first sheet only)
            df = pd.read_excel(filepath)
            
            # Verify required columns exist
            if not {'Élément', 'Montant'}.issubset(df.columns):
                raise ValueError("File must contain 'Élément' and 'Montant' columns")
            
            # Update the table
            for _, row in df.iterrows():
                element = str(row['Élément']).strip()
                amount = str(row['Montant'])  # Keep as simple string
                
                # Find matching row in table
                for table_row in range(self.input_table.rowCount()):
                    if self.input_table.item(table_row, 0).text() == element:
                        item = QTableWidgetItem(amount)
                        item.setTextAlignment(Qt.AlignCenter)
                        self.input_table.setItem(table_row, 1, item)
                        break
            
            QMessageBox.information(self, "Succès", "Données importées avec succès!")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'importation", 
                f"Le fichier doit respecter exactement ce format:\n\n"
                f"Élément,Montant\n"
                f"Résultat Net,150000\n"
                f"Dotations aux amortissements,50000\n"
                f"...\n\n"
                f"Erreur technique: {str(e)}"
            )

    
    def setup_menubar(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("Fichier")
        
        import_menu = file_menu.addMenu("Importer")
        import_csv_action = import_menu.addAction("Depuis CSV")
        import_excel_action = import_menu.addAction("Depuis Excel")
        
        export_action = file_menu.addAction("Exporter en PDF")
        file_menu.addSeparator()
        exit_action = file_menu.addAction("Quitter")
        
        # Recent files menu (will be populated dynamically)
        self.recent_files_menu = file_menu.addMenu("Fichiers récents")
        
        # Help menu
        help_menu = menubar.addMenu("Aide")
        user_guide_action = help_menu.addAction("Guide d'utilisation")
        about_action = help_menu.addAction("À propos")
        
        # Connect menu actions
        import_csv_action.triggered.connect(self.import_csv_triggered)
        import_excel_action.triggered.connect(self.import_excel_triggered)
        export_action.triggered.connect(self.export_to_pdf)
        exit_action.triggered.connect(self.close)
        user_guide_action.triggered.connect(self.show_user_guide)
        about_action.triggered.connect(self.show_about)
    
    def setup_status_bar(self):
        self.statusBar().showMessage("Prêt")
        self.calculator_status = QLabel("")
        self.statusBar().addPermanentWidget(self.calculator_status)
    
    def setup_connections(self):
        self.import_csv.clicked.connect(self.import_csv_triggered)
        self.import_excel.clicked.connect(self.import_excel_triggered)
        self.calculate_btn.clicked.connect(self.calculate)
        self.export_btn.clicked.connect(self.export_to_pdf)
    
    def load_sample_data(self):
        """Load sample data for demonstration"""
        sample_values = {
            "Résultat Net": "150000",
            "Dotations aux amortissements": "50000",
            "Variation des stocks": "-10000",
            "Variation des créances clients": "20000",
            "Variation des dettes fournisseurs": "15000",
            "Autres produits encaissables": "5000",
            "Autres charges décaissables": "10000",
            "Dividendes versés": "40000",
            "Investissements": "120000",
            "Désinvestissements": "20000"
        }
        
        for row in range(self.input_table.rowCount()):
            element = self.input_table.item(row, 0).text()
            if element in sample_values:
                self.input_table.item(row, 1).setText(sample_values[element])
                self.input_table.item(row, 1).setTextAlignment(Qt.AlignCenter)  # Center alignment

    
    def import_csv_triggered(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Importer un fichier CSV", "", "CSV Files (*.csv)"
        )
        if filepath:
            self.add_recent_file(filepath)
            try:
                df = pd.read_csv(filepath)
                self.populate_table_from_dataframe(df)
                self.update_status(f"Fichier CSV importé: {filepath}")
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Impossible de lire le fichier CSV:\n{str(e)}")
                self.update_status("Erreur d'importation CSV", is_error=True)
    
    def import_excel_triggered(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Importer un fichier Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if filepath:
            self.add_recent_file(filepath)
            try:
                df = pd.read_excel(filepath)
                self.populate_table_from_dataframe(df)
                self.update_status(f"Fichier Excel importé: {filepath}")
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Impossible de lire le fichier Excel:\n{str(e)}")
                self.update_status("Erreur d'importation Excel", is_error=True)
    
    def populate_table_from_dataframe(self, df):
        """Populate the input table from a pandas DataFrame"""
        element_mapping = {
            "resultat_net": "Résultat Net",
            "resultat": "Résultat Net",
            "net": "Résultat Net",
            "amortissements": "Dotations aux amortissements",
            "dotations": "Dotations aux amortissements",
            "stocks": "Variation des stocks",
            "creances": "Variation des créances clients",
            "dettes": "Variation des dettes fournisseurs",
            "produits": "Autres produits encaissables",
            "charges": "Autres charges décaissables",
            "dividendes": "Dividendes versés",
            "investissements": "Investissements",
            "desinvestissements": "Désinvestissements"
        }
        
        # Case insensitive column matching
        df.columns = df.columns.str.lower()
        
        for col_name in df.columns:
            normalized_name = col_name.strip().lower()
            if normalized_name in element_mapping:
                element_name = element_mapping[normalized_name]
                value = str(df[col_name].iloc[0]) if not pd.isna(df[col_name].iloc[0]) else "0"
                
                for row_idx in range(self.input_table.rowCount()):
                    if self.input_table.item(row_idx, 0).text() == element_name:
                        self.input_table.item(row_idx, 1).setText(value)
                        break
        self.input_table.viewport().update()
        print("Table updated!")  # Debug
    
    def add_recent_file(self, filepath):
        """Add file to recent files list and update menu"""
        if filepath in self.recent_files:
            self.recent_files.remove(filepath)
        self.recent_files.insert(0, filepath)
        
        # Keep only last 5 files
        self.recent_files = self.recent_files[:5]
        self.update_recent_files_menu()
    
    def update_recent_files_menu(self):
        """Update the recent files menu with current files"""
        self.recent_files_menu.clear()
        
        if not self.recent_files:
            self.recent_files_menu.setEnabled(False)
            return
        
        self.recent_files_menu.setEnabled(True)
        for filepath in self.recent_files:
            action = self.recent_files_menu.addAction(filepath)
            action.triggered.connect(lambda checked, f=filepath: self.open_recent_file(f))
    
    def open_recent_file(self, filepath):
        """Open a file from recent files list"""
        try:
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
            self.populate_table_from_dataframe(df)
            self.update_status(f"Fichier récent ouvert: {filepath}")
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Impossible d'ouvrir le fichier:\n{str(e)}")
            self.update_status("Erreur d'ouverture du fichier", is_error=True)
    
    def validate_inputs(self):
        """Validate all input values before calculation"""
        errors = []
        
        for row in range(self.input_table.rowCount()):
            item = self.input_table.item(row, 1)
            if not item or not item.text():
                errors.append(f"Valeur manquante pour {self.input_table.item(row, 0).text()}")
                continue
                
            try:
                value = float(item.text())
                if value < 0 and row not in [2,3,4]:  # Only some fields can be negative
                    errors.append(f"Valeur négative non autorisée pour {self.input_table.item(row, 0).text()}")
            except ValueError:
                errors.append(f"Valeur invalide pour {self.input_table.item(row, 0).text()}")
        
        if errors:
            QMessageBox.warning(self, "Erreurs de validation", "\n".join(errors))
            self.update_status("Erreurs dans les données d'entrée", is_error=True)
            return False
        return True
    
    def calculate(self):
        """Perform all financial calculations"""
        if not self.validate_inputs():
            return
            
        try:
            # Get values from table
            values = {}
            for row in range(self.input_table.rowCount()):
                item = self.input_table.item(row, 1)
                values[self.input_table.item(row, 0).text()] = float(item.text()) if item and item.text() else 0
            
            # Calculate CAF
            caf = (values["Résultat Net"] + 
                  values["Dotations aux amortissements"] +
                  values["Variation des stocks"] +
                  values["Variation des créances clients"] +
                  values["Variation des dettes fournisseurs"] +
                  values["Autres produits encaissables"] -
                  values["Autres charges décaissables"])
            
            # Calculate Autofinancement
            autofinancement = caf - values["Dividendes versés"]
            
            # Calculate Taux d'autofinancement
            investissements_net = values["Investissements"] - values["Désinvestissements"]
            taux = (autofinancement / investissements_net) * 100 if investissements_net != 0 else 0
            
            # Update UI
            self.caf_value.setText(f"{caf:,.2f} €")
            self.autofinancement_value.setText(f"{autofinancement:,.2f} €")
            self.taux_value.setText(f"{taux:,.2f} %")
            
            # Update interpretation
            self.interpretation_value.setText(self.get_interpretation(caf, autofinancement, taux))
            
            # Update chart
            self.update_chart(caf, autofinancement, taux)
            
            self.update_status("Calcul terminé avec succès")
            
        except Exception as e:
            QMessageBox.warning(self, "Erreur de calcul", f"Une erreur est survenue:\n{str(e)}")
            self.update_status("Erreur lors du calcul", is_error=True)
    
    def get_interpretation(self, caf, autofinancement, taux):
        """Generate detailed financial interpretation"""
        interpretations = []
        
        # CAF analysis
        if caf > 0:
            interpretations.append(
                f"L'entreprise dégage une capacité d'autofinancement positive de {caf:,.2f}€, "
                "indiquant qu'elle génère des ressources internes suffisantes."
            )
        else:
            interpretations.append(
                f"Alerte: La CAF est négative ({caf:,.2f}€), ce qui signifie que l'entreprise "
                "ne génère pas suffisamment de ressources internes."
            )
        
        # Autofinancement analysis
        if autofinancement > caf * 0.7:
            interpretations.append(
                "La politique de dividendes est conservative, préservant les ressources "
                "pour l'entreprise."
            )
        elif autofinancement > 0:
            interpretations.append(
                "Une partie significative des ressources est distribuée aux actionnaires."
            )
        else:
            interpretations.append(
                "Alerte: L'autofinancement est négatif, les dividendes dépassent la CAF."
            )
        
        # Taux analysis
        if taux > 100:
            interpretations.append(
                f"Excellent taux d'autofinancement ({taux:.2f}%). L'entreprise finance "
                "intégralement ses investissements et dégage un excédent."
            )
        elif taux > 80:
            interpretations.append(
                f"Bon taux d'autofinancement ({taux:.2f}%). L'entreprise finance la grande "
                "majorité de ses investissements par ses ressources internes."
            )
        elif taux > 50:
            interpretations.append(
                f"Taux d'autofinancement modéré ({taux:.2f}%). L'entreprise combine "
                "financement interne et externe."
            )
        elif taux > 0:
            interpretations.append(
                f"Taux d'autofinancement faible ({taux:.2f}%). Forte dépendance aux "
                "financements externes."
            )
        else:
            interpretations.append(
                f"Alerte: Taux d'autofinancement critique ({taux:.2f}%). Situation financière "
                "très fragile."
            )
        
        return "\n\n".join(interpretations)
    
    def update_chart(self, caf, autofinancement, taux):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        # Configuration de l'espacement
        self.figure.subplots_adjust(left=0.15, right=0.95, top=0.9, bottom=0.2)
        
        # Données pour le graphique
        labels = ['CAF', 'Autofinancement']
        values = [caf, autofinancement]
        colors = ['#4CAF50', '#2196F3']
        
        # Création des barres
        bars = ax.bar(labels, values, color=colors)
        
        # Formatage des valeurs
        ax.yaxis.set_major_formatter('€{x:,.0f}')
        
        # Ajout des valeurs sur les barres
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
            f'€{height:,.0f}',
            ha='center', va='bottom',
            fontsize=10)
        
        # Configuration des axes
        ax.set_ylim(0, max(values)*1.2)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        # Titre et informations
        #ax.set_title('Analyse d\'Autofinancement', pad=20)
        #ax.text(0.5, 1.1, f'Taux: {taux:.2f}%', transform=ax.transAxes, ha='left', bbox=dict(facecolor='white', edgecolor='lightgray', boxstyle='round'))
        
        self.canvas.draw()
        """Update the chart visualization"""
        self.figure.clear()
        ax = self.figure.add_subplot(111)
            
            # Bar chart for CAF and Autofinancement
        labels = ['CAF', 'Autofinancement']
        values = [caf, autofinancement]
        bars = ax.bar(labels, values, color=['#2196F3', '#4CAF50'])
            
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                    f'{height:,.2f}€',
                    ha='center', va='bottom')
            
            # Add taux as text
            #ax.text(0.5, 0.95, f'Taux d\'autofinancement: {taux:.2f}%', transform=ax.transAxes, ha='center', va='top', bbox=dict(facecolor='white', alpha=0.8))
            
            ax.set_ylabel('Montant (€)')
            #ax.set_title('Analyse d\'Autofinancement')
            self.canvas.draw()




    def export_to_pdf(self):
        """Export the current results to PDF"""
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Exporter en PDF", "", "PDF Files (*.pdf)"
        )
        if filepath:
            try:
                # Save chart temporarily
                chart_path = "temp_chart.png"
                self.figure.savefig(chart_path, bbox_inches='tight', dpi=150)
                
                # Create PDF
                printer = QPrinter(QPrinter.HighResolution)
                printer.setOutputFormat(QPrinter.PdfFormat)
                printer.setOutputFileName(filepath)
                
                # Create document
                doc = QTextDocument()
                html = self.generate_report_html(chart_path)
                doc.setHtml(html)
                
                # Print to PDF
                doc.print_(printer)
                
                QMessageBox.information(self, "Succès", "Rapport exporté avec succès!")
                self.update_status(f"Rapport exporté: {filepath}")
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Échec de l'export:\n{str(e)}")
                self.update_status("Erreur lors de l'export PDF", is_error=True)
    
    def generate_report_html(self, chart_path):
        """Generate HTML content for the PDF report"""
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
                <span>Capacité d'Autofinancement (CAF): </span>
                <span class="value">{self.caf_value.text()}</span>
            </div>
            <div class="result">
                <span>Autofinancement: </span>
                <span class="value">{self.autofinancement_value.text()}</span>
            </div>
            <div class="result">
                <span>Taux d'Autofinancement: </span>
                <span class="value">{self.taux_value.text()}</span>
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
    
    def show_user_guide(self):
        """Show the user guide dialog"""
        help_text = """
        <h2>Guide d'Utilisation du Calculateur d'Autofinancement</h2>
        
        <h3>1. Saisie des Données</h3>
        <p>Vous pouvez saisir les données manuellement dans le tableau ou les importer depuis un fichier CSV/Excel.</p>
        
        <h3>2. Éléments Requis</h3>
        <ul>
            <li><b>Résultat Net</b>: Le bénéfice net de l'entreprise après impôts</li>
            <li><b>Dotations aux amortissements</b>: Charges non décaissables comptabilisées</li>
            <li><b>Variation des stocks</b>: Différence entre stocks finaux et initiaux</li>
            <li><b>Variation des créances clients</b>: Variation des créances clients et autres créances</li>
            <li><b>Variation des dettes fournisseurs</b>: Variation des dettes fournisseurs et autres dettes</li>
            <li><b>Autres produits encaissables</b>: Autres produits qui se traduiront par des encaissements</li>
            <li><b>Autres charges décaissables</b>: Autres charges qui se traduiront par des décaissements</li>
            <li><b>Dividendes versés</b>: Dividendes distribués aux actionnaires</li>
            <li><b>Investissements</b>: Total des investissements réalisés</li>
            <li><b>Désinvestissements</b>: Produits des cessions d'actifs</li>
        </ul>
        
        <h3>3. Calcul et Résultats</h3>
        <p>Cliquez sur le bouton "Calculer" pour obtenir :</p>
        <ul>
            <li>La Capacité d'Autofinancement (CAF)</li>
            <li>Le montant d'Autofinancement</li>
            <li>Le taux d'Autofinancement</li>
            <li>Une interprétation automatique des résultats</li>
            <li>Une visualisation graphique</li>
        </ul>
        
        <h3>4. Export PDF</h3>
        <p>Générez un rapport professionnel au format PDF contenant tous les résultats.</p>
        """
        
        dialog = QMessageBox(self)
        dialog.setWindowTitle("Guide d'Utilisation")
        dialog.setTextFormat(Qt.RichText)
        dialog.setText(help_text)
        dialog.exec()
    
    def show_about(self):
        """Show the about dialog"""
        about_text = """
        <center>
        <h1>Calculateur d'Autofinancement</h1>
        <p>Version 1.0</p>
        <br>
        <p>Cet outil permet de calculer et analyser :</p>
        <ul style="text-align: left">
            <li>La Capacité d'Autofinancement (CAF)</li>
            <li>L'Autofinancement</li>
            <li>Le taux d'Autofinancement</li>
        </ul>
        <br>
        <p>Développé par :</p>
        <p>NEDJAR Nassim & DJAFERCHERIF Soumia<br>    
        <br>
        <p>© 2025 Tous droits réservés</p>
        </center>
        """
        QMessageBox.about(self, "À propos", about_text)
    
    def update_status(self, message, is_error=False):
        """Update the status bar with a message"""
        self.statusBar().showMessage(message)
        if is_error:
            self.calculator_status.setText("Erreur")
            self.calculator_status.setStyleSheet("color: red;")
        else:
            self.calculator_status.setText("OK")
            self.calculator_status.setStyleSheet("color: green;")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Load the stylesheet
    with open("style.css", "r") as f:
        app.setStyleSheet(f.read())
    
    calculator = AutofinancementCalculator()
    calculator.show()
    sys.exit(app.exec())