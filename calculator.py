import sys
from datetime import datetime
import traceback
import pandas as pd
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QScrollArea, QHeaderView,
    QSizePolicy, QLineEdit
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QTextDocument, QDoubleValidator
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
        
        # Dividend input section
        dividend_frame = QFrame()
        dividend_layout = QHBoxLayout(dividend_frame)
        dividend_layout.setContentsMargins(0, 10, 0, 10)
        
        dividend_label = QLabel("Dividendes (Compte 457):")
        dividend_label.setObjectName("dividendLabel")
        dividend_label.setStyleSheet("font-size: 14px; padding: 5px;") 

        
        self.dividend_input = QLineEdit()
        self.dividend_input.setObjectName("dividendInput")
        self.dividend_input.setStyleSheet("font-size: 14px; padding: 5px;") 
        self.dividend_input.setPlaceholderText("Entrez le montant...")
        self.dividend_input.setValidator(QDoubleValidator())
        
        dividend_layout.addWidget(dividend_label)
        dividend_layout.addWidget(self.dividend_input, stretch=1)
        main_layout.addWidget(dividend_frame)
        
        # Calculate button
        self.calculate_btn = QPushButton("Calculer")
        self.calculate_btn.setObjectName("calculateButton")
        self.calculate_btn.setFixedSize(200, 50)
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

        # Button row
        button_row = QHBoxLayout()
        button_row.setSpacing(15)
        
        self.import_excel = QPushButton("Importer Excel")
        self.import_excel.setObjectName("importButton")
        self.import_excel.setFixedHeight(40)
        self.import_excel.clicked.connect(self.import_xlsx_data)
        
        self.add_row_button = QPushButton("+ Ajouter une ligne")
        self.add_row_button.setObjectName("addRowButton")
        self.add_row_button.setFixedHeight(40)
        self.add_row_button.clicked.connect(self.add_table_row)
        
        button_row.addWidget(self.import_excel)
        button_row.addWidget(self.add_row_button)
        layout.addLayout(button_row)

        # Table
        table_scroll = QScrollArea()
        table_scroll.setWidgetResizable(True)

        self.input_table = QTableWidget()
        self.setup_table()

        table_scroll.setWidget(self.input_table)
        layout.addWidget(table_scroll)

        return frame
    
    def add_table_row(self):
        row_position = self.input_table.rowCount()
        self.input_table.insertRow(row_position)
        self.input_table.setItem(row_position, 0, QTableWidgetItem(""))
        self.input_table.setItem(row_position, 1, QTableWidgetItem(""))
        self.input_table.setItem(row_position, 2, QTableWidgetItem("0"))

    def setup_table(self):
        self.input_table.setColumnCount(3)
        self.input_table.setHorizontalHeaderLabels(["Libellé", "Compte", "Montant (DZD)"])
        self.input_table.setRowCount(0)

        self.input_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.input_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.input_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.input_table.verticalHeader().setDefaultSectionSize(40)
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
            # Lire le fichier Excel
            try:
                df = pd.read_excel(filepath, engine='openpyxl')
            except:
                df = pd.read_excel(filepath, engine='xlrd')

            # Normaliser les noms de colonnes
            df.columns = [str(col).strip().lower() for col in df.columns]
            
            # Détection des colonnes
            compte_col = next((col for col in df.columns if 'compte' in col or 'numéro' in col or 'numero' in col), None)
            montant_col = next((col for col in df.columns if 'montant' in col or 'valeur' in col), None)
            libelle_col = next((col for col in df.columns if 'libellé' in col or 'libelle' in col or 'désignation' in col), None)

            if not compte_col or not montant_col:
                raise ValueError("Colonnes requises non trouvées: besoin d'une colonne 'compte' et 'montant'")

            # Filtrer les lignes valides (où le compte n'est pas NaN et est un string non vide)
            df = df[pd.notna(df[compte_col]) & (df[compte_col].astype(str).str.strip() != '')].copy()
            
            # Vider le tableau avant l'import
            self.input_table.setRowCount(0)

            # Traitement des lignes valides
            valid_rows = 0
            for _, row in df.iterrows():
                try:
                    compte = str(row[compte_col]).strip()
                    if not compte:  # Si le compte est une string vide
                        continue
                        
                    libelle = str(row[libelle_col]).strip() if libelle_col and pd.notna(row.get(libelle_col)) else ""
                    
                    try:
                        montant = float(str(row[montant_col]).replace(',', '')) if pd.notna(row[montant_col]) else 0.0
                    except ValueError:
                        montant = 0.0

                    current_row = self.input_table.rowCount()
                    self.input_table.insertRow(current_row)
                    self.input_table.setItem(current_row, 0, QTableWidgetItem(libelle))
                    self.input_table.setItem(current_row, 1, QTableWidgetItem(compte))
                    self.input_table.setItem(current_row, 2, QTableWidgetItem(f"{montant:,.2f}"))
                    valid_rows += 1

                except Exception as e:
                    print(f"Ignoring row due to error: {e}")
                    continue

            QMessageBox.information(
                self, 
                "Succès", 
                f"Import terminé.\n"
                f"- Lignes valides importées: {valid_rows}"
            )
            
            if valid_rows > 0:
                self.calculate()

        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'importation :\n{str(e)}")
    
    def calculate(self):
        try:
            # Initialisation des variables
            resultat_net = 0.0
            dotations = 0.0
            valeur_cession = 0.0
            reprises = 0.0
            produits_cession = 0.0
            subventions = 0.0

            # Fonction de conversion
            def convert_montant(text):
                try:
                    text = str(text).strip().replace(' ', '').replace(',', '')
                    return float(text) if text else 0.0
                except:
                    return 0.0

            # Parcours du tableau
            for row in range(self.input_table.rowCount()):
                compte_item = self.input_table.item(row, 1)
                montant_item = self.input_table.item(row, 2)

                if not compte_item or not montant_item:
                    continue

                compte = compte_item.text().strip()
                montant = convert_montant(montant_item.text())

                # Logique de calcul
                if compte.startswith('12'):
                    resultat_net += montant
                elif any(compte.startswith(c) for c in ['681', '686', '687']):
                    dotations += montant
                elif compte.startswith('675'):
                    valeur_cession += montant
                elif any(compte.startswith(c) for c in ['781', '786', '787']):
                    reprises += montant
                elif compte.startswith('775'):
                    produits_cession += montant
                elif compte.startswith('777'):
                    subventions += montant

            # Récupération des dividendes depuis le champ input
            try:
                dividendes = convert_montant(self.dividend_input.text())
            except:
                dividendes = 0.0

            # Calculs finaux
            caf = (resultat_net + dotations + valeur_cession - reprises - produits_cession - subventions)
            autofinancement = caf - dividendes

            # Formatage de l'affichage
            def format_montant(value):
                return f"{value:,.0f} DZD".replace(",", " ") if value == int(value) else f"{value:,.2f} DZD".replace(",", " ")

            # Mise à jour de l'interface
            self.resultat_net_value.setText(format_montant(resultat_net))
            self.caf_value.setText(format_montant(caf))
            self.autofinancement_value.setText(format_montant(autofinancement))
            self.interpretation_value.setText(self.get_interpretation(resultat_net, caf, autofinancement, dividendes))
            
            self.update_chart(resultat_net, caf, autofinancement)

        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur de calcul:\n{str(e)}")
            print(f"Erreur complète:\n{traceback.format_exc()}")
    

    def get_interpretation(self, resultat_net, caf, autofinancement, dividendes):
        interpretations = []
        
        # Analyse du résultat net
        if resultat_net > 0:
            interpretations.append(
                "🔹 <b>Résultat Net Positif</b> ({} DZD):\n"
                "L'entreprise dégage un bénéfice. Cela indique que les produits dépassent les charges, "
                "ce qui est un signe de bonne santé financière à court terme.".format(f"{resultat_net:,.2f}")
            )
        else:
            interpretations.append(
                "🔹 <b>Résultat Net Négatif</b> ({} DZD):\n"
                "L'entreprise est en situation de perte. Cela peut être préoccupant si la tendance persiste, "
                "mais peut être normal pour une entreprise en phase d'investissement ou de démarrage.".format(f"{resultat_net:,.2f}")
            )
        
        # Analyse de la CAF
        if caf > resultat_net:
            cash_info = "La CAF est supérieure au résultat net, ce qui est normal car elle inclut les dotations non décaissables."
        elif caf == resultat_net:
            cash_info = "La CAF est égale au résultat net, ce qui est rare et peut indiquer l'absence de dotations."
        else:
            cash_info = "La CAF est inférieure au résultat net, situation atypique qui mérite investigation."
        
        if caf > 0:
            interpretations.append(
                "🔹 <b>CAF Positive</b> ({} DZD):\n"
                "L'entreprise génère des liquidités internes suffisantes pour:\n"
                "- Financer ses investissements\n"
                "- Rembourser ses dettes\n"
                "- Payer des dividendes\n{}".format(f"{caf:,.2f}", cash_info)
            )
        else:
            interpretations.append(
                "🔹 <b>CAF Négative</b> ({} DZD):\n"
                "Attention! L'entreprise ne génère pas assez de cash flow interne.\n"
                "Cela peut entraîner:\n"
                "- Des difficultés de trésorerie\n"
                "- Une dépendance accrue au financement externe\n"
                "- Des risques de cessation de paiement".format(f"{caf:,.2f}")
            )
        
        # Analyse de l'autofinancement
        if autofinancement > 0:
            if dividendes > 0:
                dividend_info = (
                    "L'entreprise peut à la fois:\n"
                    "- Financer sa croissance ({:,.2f} DZD disponibles)\n"
                    "- Récompenser ses actionnaires ({:,.2f} DZD de dividendes)".format(autofinancement, dividendes)
                )
            else:
                dividend_info = "L'entreprise conserve toutes ses ressources pour financer son développement."
            
            interpretations.append(
                "🔹 <b>Autofinancement Positif</b> ({} DZD):\n"
                "Situation très favorable. {}\n"
                "La politique de dividendes semble soutenable.".format(f"{autofinancement:,.2f}", dividend_info)
            )
        else:
            interpretations.append(
                "🔹 <b>Autofinancement Négatif</b> ({} DZD):\n"
                "Situation risquée! L'entreprise distribue plus de dividendes ({:,.2f} DZD) "
                "qu'elle ne génère de CAF.\n"
                "Cela peut conduire à:\n"
                "- Un endettement excessif\n"
                "- Une réduction des investissements\n"
                "- A terme, une baisse de compétitivité".format(f"{autofinancement:,.2f}", dividendes)
            )
        
        # Recommandations globales
        recommendations = []
        if autofinancement < 0:
            recommendations.append(
                "🚩 <b>Recommandation urgente</b>: Réviser la politique de dividendes à la baisse "
                "ou trouver des sources de financement externes."
            )
        elif caf < resultat_net:
            recommendations.append(
                "🔍 <b>Vérifier</b>: La composition de la CAF pour comprendre pourquoi elle est inférieure au résultat net."
            )
        
        if recommendations:
            interpretations.append("\n<b>RECOMMANDATIONS:</b>\n" + "\n".join(recommendations))
        
        return "<br>".join(interpretations).replace("\n", "<br>")

    
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