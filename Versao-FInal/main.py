# main.py (VERSÃO ATUALIZADA E COMPLETA)

from openpyxl.styles import PatternFill, Font
from openpyxl import load_workbook
import sys
import os
import json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox,
                             QFormLayout, QLineEdit, QComboBox, QTableWidget, 
                             QTableWidgetItem, QDialog, QInputDialog, QHeaderView)
from PyQt5.QtCore import Qt

# <<< IMPORTAÇÕES DAS CLASSES ENCAPSULADAS >>>
from code_manager import CodeGenerator
from history_manager import HistoryManager
from history_dialog import HistoryDialog
from processing import ProcessThread
from nesting_dialog import NestingDialog # Importa a nova classe do diálogo de nesting
from calculo_cortes import calcular_plano_de_corte

# =============================================================================
# ESTILO VISUAL DA APLICAÇÃO (QSS - Qt StyleSheet)
# =============================================================================
INOVA_PROCESS_STYLE = """
/* ================================================================================
    Estilo Dark Theme para INOVA PROCESS
    - Paleta de cores baseada em tons de azul, ciano e cinza para um visual
      tecnológico e profissional.
    - Foco em legibilidade e usabilidade dos componentes.
================================================================================
*/

/* Estilo Geral da Janela e Widgets */
QWidget {
    background-color: #2D3748; /* Azul-acinzentado escuro para o fundo */
    color: #E2E8F0; /* Texto em cinza claro para alto contraste */
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 10pt;
    border: none;
}

/* Estilo para GroupBox (Contêineres com Título) */
QGroupBox {
    background-color: #1A202C; /* Fundo mais escuro para destaque */
    border: 1px solid #4A5568;
    border-radius: 8px;
    margin-top: 1em; /* Espaço para o título não sobrepor a borda */
    font-weight: bold;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 2px 8px;
    background-color: #00B5D8; /* Ciano vibrante para o título */
    color: #1A202C; /* Texto escuro no título para contraste */
    border-radius: 4px;
}

/* Campos de Entrada de Texto, ComboBox e SpinBox */
QLineEdit, QTextEdit, QComboBox, QDoubleSpinBox, QSpinBox {
    background-color: #2D3748;
    border: 1px solid #4A5568;
    border-radius: 4px;
    padding: 5px;
    color: #E2E8F0;
}

QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QSpinBox:focus {
    border: 1px solid #00B5D8; /* Destaque em ciano ao focar */
}

/* Estilo para ComboBox (Dropdown) */
QComboBox::drop-down {
    border: none;
}

QComboBox::down-arrow {
    image: url(C:/Users/mathe/Desktop/INOVA_PROCESS/down_arrow.png); /* Use um ícone de seta branca aqui */
    width: 12px;
    height: 12px;
    margin-right: 8px;
}

QComboBox QAbstractItemView {
    background-color: #2D3748;
    border: 1px solid #00B5D8;
    selection-background-color: #00B5D8;
    selection-color: #1A202C;
    outline: 0px;
}

/* Botões */
QPushButton {
    background-color: #4A5568;
    color: #E2E8F0;
    font-weight: bold;
    padding: 8px 12px;
    border-radius: 4px;
}

QPushButton:hover {
    background-color: #718096; /* Efeito hover mais claro */
}

QPushButton:pressed {
    background-color: #2D3748;
}

/* Botões de Ação Principal (Ex: Gerar, Calcular) */
QPushButton#primaryButton {
    background-color: #00B5D8;
    color: #1A202C;
}
QPushButton#primaryButton:hover {
    background-color: #4FD1C5; /* Verde-água para hover */
}
QPushButton#primaryButton:pressed {
    background-color: #00A3BF;
}

/* <<< ADICIONADO: Estilo para botão de sucesso (verde) >>> */
QPushButton#successButton {
    background-color: #107C10; /* Verde escuro */
    color: #FFFFFF; /* Texto branco */
}
QPushButton#successButton:hover {
    background-color: #159d15;
}
QPushButton#successButton:pressed {
    background-color: #0c5a0c;
}

/* <<< ADICIONADO: Estilo para botão de aviso (amarelo) >>> */
QPushButton#warningButton {
    background-color: #DCA307; /* Amarelo/Ouro */
    color: #1A202C; /* Texto escuro */
}
QPushButton#warningButton:hover {
    background-color: #f0b92a;
}
QPushButton#warningButton:pressed {
    background-color: #c49106;
}


/* ============================================================================== */
/* =================== CORREÇÃO PRINCIPAL: TABELA E LISTAS ====================== */
/* ============================================================================== */

QTableWidget, QListView {
    background-color: #1A202C; /* Fundo da área da tabela */
    border: 1px solid #4A5568;
    border-radius: 4px;
    gridline-color: #4A5568; /* Cor da grade */
}

/* Cabeçalho da Tabela */
QHeaderView::section {
    background-color: #2D3748;
    color: #E2E8F0;
    padding: 6px;
    border: 1px solid #4A5568;
    font-weight: bold;
}

/* Itens da Tabela - Garante que o texto seja visível */
QTableWidget::item {
    color: #E2E8F0;
    font-size: 11pt; /* <<< AUMENTA A FONTE PARA MELHOR LEITURA >>> */
    padding-left: 5px;
}

/* Item da Tabela quando Selecionado */
QTableWidget::item:selected {
    background-color: #00B5D8; /* Fundo ciano na seleção */
    color: #1A202C; /* Texto escuro para contraste na seleção */
}

/* Barra de Log/Execução */
QTextEdit#logExecution {
    font-family: 'Courier New', Courier, monospace;
    background-color: #1A202C;
    color: #4FD1C5; /* Texto verde-água, estilo "terminal" */
}

/* Barra de Rolagem */
QScrollBar:vertical {
    border: none;
    background: #1A202C;
    width: 12px;
    margin: 0px 0px 0px 0px;
}

QScrollBar::handle:vertical {
    background: #4A5568;
    min-height: 20px;
    border-radius: 6px;
}
QScrollBar::handle:vertical:hover {
    background: #718096;
}

QScrollBar:horizontal {
    border: none;
    background: #1A202C;
    height: 12px;
    margin: 0px 0px 0px 0px;
}

QScrollBar::handle:horizontal {
    background: #4A5568;
    min-width: 20px;
    border-radius: 6px;
}

QScrollBar::handle:horizontal:hover {
    background: #718096;
}

QScrollBar::add-line, QScrollBar::sub-line {
    border: none;
    background: none;
}
"""

# =============================================================================
# CLASSE PRINCIPAL DA INTERFACE GRÁFICA
# =============================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Desenhos Técnicos e DXF")
        self.setGeometry(100, 100, 1100, 850)
        
        # Instancia as classes dos módulos importados
        self.code_generator = CodeGenerator()
        self.history_manager = HistoryManager()
        
        # Variáveis de estado da aplicação
        self.colunas_df = ['nome_arquivo', 'forma', 'espessura', 'qtd', 'largura', 'altura', 'diametro', 'rt_base', 'rt_height', 'trapezoid_large_base', 'trapezoid_small_base', 'trapezoid_height', 'furos']
        self.manual_df = pd.DataFrame(columns=self.colunas_df)
        self.excel_df = pd.DataFrame(columns=self.colunas_df)
        self.furos_atuais = []
        self.project_directory = None

        # =====================================================================
        # <<< INÍCIO DA CONSTRUÇÃO DA INTERFACE GRÁFICA (UI) >>>
        # =====================================================================
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        top_h_layout = QHBoxLayout()
        left_v_layout = QVBoxLayout()
        
        # --- Grupo 1: Projeto ---
        project_group = QGroupBox("1. Projeto")
        project_layout = QVBoxLayout()
        self.start_project_btn = QPushButton("Iniciar Novo Projeto...")
        self.history_btn = QPushButton("Ver Histórico de Projetos")
        project_layout.addWidget(self.start_project_btn)
        project_layout.addWidget(self.history_btn)
        project_group.setLayout(project_layout)
        left_v_layout.addWidget(project_group)
        
        # --- Grupo 2: Carregar Planilha ---
        file_group = QGroupBox("2. Carregar Planilha (Opcional)")
        file_layout = QVBoxLayout()
        self.file_label = QLabel("Nenhum projeto ativo.")
        file_button_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("Selecionar Planilha")
        self.clear_excel_btn = QPushButton("Limpar Planilha")
        file_button_layout.addWidget(self.select_file_btn)
        file_button_layout.addWidget(self.clear_excel_btn)
        file_layout.addWidget(self.file_label)
        file_layout.addLayout(file_button_layout)
        file_group.setLayout(file_layout)
        left_v_layout.addWidget(file_group)

        # --- Grupo 3: Informações da Peça ---
        manual_group = QGroupBox("3. Informações da Peça")
        manual_layout = QFormLayout()
        self.projeto_input = QLineEdit()
        self.projeto_input.setReadOnly(True)
        manual_layout.addRow("Nº do Projeto Ativo:", self.projeto_input)
        self.nome_input = QLineEdit()
        self.generate_code_btn = QPushButton("Gerar Código")
        name_layout = QHBoxLayout()
        name_layout.addWidget(self.nome_input)
        name_layout.addWidget(self.generate_code_btn)
        manual_layout.addRow("Nome/ID da Peça:", name_layout)
        self.forma_combo = QComboBox()
        self.forma_combo.addItems(['rectangle', 'circle', 'right_triangle', 'trapezoid'])
        self.espessura_input, self.qtd_input = QLineEdit(), QLineEdit()
        manual_layout.addRow("Forma:", self.forma_combo)
        manual_layout.addRow("Espessura (mm):", self.espessura_input)
        manual_layout.addRow("Quantidade:", self.qtd_input)
        self.largura_input, self.altura_input = QLineEdit(), QLineEdit()
        self.diametro_input, self.rt_base_input, self.rt_height_input = QLineEdit(), QLineEdit(), QLineEdit()
        self.trapezoid_large_base_input, self.trapezoid_small_base_input, self.trapezoid_height_input = QLineEdit(), QLineEdit(), QLineEdit()
        self.largura_row = [QLabel("Largura:"), self.largura_input]; manual_layout.addRow(*self.largura_row)
        self.altura_row = [QLabel("Altura:"), self.altura_input]; manual_layout.addRow(*self.altura_row)
        self.diametro_row = [QLabel("Diâmetro:"), self.diametro_input]; manual_layout.addRow(*self.diametro_row)
        self.rt_base_row = [QLabel("Base Triângulo:"), self.rt_base_input]; manual_layout.addRow(*self.rt_base_row)
        self.rt_height_row = [QLabel("Altura Triângulo:"), self.rt_height_input]; manual_layout.addRow(*self.rt_height_row)
        self.trap_large_base_row = [QLabel("Base Maior:"), self.trapezoid_large_base_input]; manual_layout.addRow(*self.trap_large_base_row)
        self.trap_small_base_row = [QLabel("Base Menor:"), self.trapezoid_small_base_input]; manual_layout.addRow(*self.trap_small_base_row)
        self.trap_height_row = [QLabel("Altura:"), self.trapezoid_height_input]; manual_layout.addRow(*self.trap_height_row)
        manual_group.setLayout(manual_layout)
        left_v_layout.addWidget(manual_group)
        left_v_layout.addStretch()
        top_h_layout.addLayout(left_v_layout)
        
        # --- Grupo 4: Furos ---
        furos_main_group = QGroupBox("4. Adicionar Furos")
        furos_main_layout = QVBoxLayout()
        self.rep_group = QGroupBox("Furação Rápida")
        rep_layout = QFormLayout()
        self.rep_diam_input, self.rep_offset_input = QLineEdit(), QLineEdit()
        rep_layout.addRow("Diâmetro Furos:", self.rep_diam_input)
        rep_layout.addRow("Offset Borda:", self.rep_offset_input)
        self.replicate_btn = QPushButton("Replicar Furos")
        rep_layout.addRow(self.replicate_btn)
        self.rep_group.setLayout(rep_layout)
        furos_main_layout.addWidget(self.rep_group)
        man_group = QGroupBox("Furos Manuais")
        man_layout = QVBoxLayout()
        man_form_layout = QFormLayout()
        self.diametro_furo_input, self.pos_x_input, self.pos_y_input = QLineEdit(), QLineEdit(), QLineEdit()
        man_form_layout.addRow("Diâmetro:", self.diametro_furo_input)
        man_form_layout.addRow("Posição X:", self.pos_x_input)
        man_form_layout.addRow("Posição Y:", self.pos_y_input)
        self.add_furo_btn = QPushButton("Adicionar Furo Manual")
        man_layout.addLayout(man_form_layout)
        man_layout.addWidget(self.add_furo_btn)
        self.furos_table = QTableWidget(0, 4)
        self.furos_table.setMaximumHeight(150) # <<< LIMITA A ALTURA DA TABELA DE FUROS
        self.furos_table.setHorizontalHeaderLabels(["Diâmetro", "Pos X", "Pos Y", "Ação"])
        man_layout.addWidget(self.furos_table)
        man_group.setLayout(man_layout)
        furos_main_layout.addWidget(man_group)
        furos_main_group.setLayout(furos_main_layout)
        top_h_layout.addWidget(furos_main_group, stretch=1)
        main_layout.addLayout(top_h_layout)
        
        self.add_piece_btn = QPushButton("Adicionar Peça à Lista")
        main_layout.addWidget(self.add_piece_btn)

        # --- Grupo 5: Lista de Peças ---
        list_group = QGroupBox("5. Lista de Peças para Produção")
        list_layout = QVBoxLayout()
        self.pieces_table = QTableWidget()
        self.table_headers = [col.replace('_', ' ').title() for col in self.colunas_df] + ["Ações"]
        self.pieces_table.setColumnCount(len(self.table_headers))
        self.pieces_table.setHorizontalHeaderLabels(self.table_headers)
        list_layout.addWidget(self.pieces_table)
        self.dir_label = QLabel("Nenhum projeto ativo. Inicie um novo projeto.")
        self.dir_label.setStyleSheet("font-style: italic; color: grey;")
        list_layout.addWidget(self.dir_label)
        process_buttons_layout = QHBoxLayout()
        self.conclude_project_btn = QPushButton("Projeto Concluído")
        self.export_excel_btn = QPushButton("Exportar para Excel")
        self.process_pdf_btn, self.process_dxf_btn, self.process_all_btn = QPushButton("Gerar PDFs"), QPushButton("Gerar DXFs"), QPushButton("Gerar PDFs e DXFs")
        process_buttons_layout.addWidget(self.export_excel_btn)
        process_buttons_layout.addWidget(self.conclude_project_btn)
        process_buttons_layout.addStretch()
        self.calculate_nesting_btn = QPushButton("Calcular Aproveitamento")
        process_buttons_layout.addWidget(self.calculate_nesting_btn)
        process_buttons_layout.addWidget(self.process_pdf_btn)
        process_buttons_layout.addWidget(self.process_dxf_btn)
        process_buttons_layout.addWidget(self.process_all_btn)
        list_layout.addLayout(process_buttons_layout)
        list_group.setLayout(list_layout)
        main_layout.addWidget(list_group, stretch=5) # <<< AUMENTA A PRIORIDADE DE EXPANSÃO DA LISTA
        
        # --- Barra de Progresso e Log ---
        self.progress_bar = QProgressBar()
        main_layout.addWidget(self.progress_bar)
        log_group = QGroupBox("Log de Execução")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group, stretch=1) # <<< DÁ UMA PRIORIDADE MENOR AO LOG

        # --- Aplicação de Estilos Específicos via objectName ---
        self.start_project_btn.setObjectName("primaryButton")
        self.conclude_project_btn.setObjectName("successButton")
        self.calculate_nesting_btn.setObjectName("warningButton")

        self.statusBar().showMessage("Pronto")
        
        # --- Conexões de Sinais e Slots (Eventos) ---
        self.calculate_nesting_btn.clicked.connect(self.open_nesting_dialog)
        self.start_project_btn.clicked.connect(self.start_new_project)
        self.history_btn.clicked.connect(self.show_history_dialog)
        self.select_file_btn.clicked.connect(self.select_file)
        self.clear_excel_btn.clicked.connect(self.clear_excel_data)
        self.generate_code_btn.clicked.connect(self.generate_piece_code)
        self.add_piece_btn.clicked.connect(self.add_manual_piece)
        self.forma_combo.currentTextChanged.connect(self.update_dimension_fields)
        self.replicate_btn.clicked.connect(self.replicate_holes)
        self.add_furo_btn.clicked.connect(self.add_furo_temp)
        self.process_pdf_btn.clicked.connect(self.start_pdf_generation)
        self.process_dxf_btn.clicked.connect(self.start_dxf_generation)
        self.process_all_btn.clicked.connect(self.start_all_generation)
        self.conclude_project_btn.clicked.connect(self.conclude_project)
        self.export_excel_btn.clicked.connect(self.export_project_to_excel)
        
        # --- Estado Inicial da UI ---
        self.set_initial_button_state()
        self.update_dimension_fields(self.forma_combo.currentText())

    # =====================================================================
    # <<< INÍCIO DAS FUNÇÕES (MÉTODOS) DA CLASSE MainWindow >>>
    # =====================================================================
    
    # ... (TODOS OS SEUS MÉTODOS CONTINUAM IGUAIS ATÉ O update_table_display)
    # Exemplo: start_new_project, set_initial_button_state, etc.
    # Vou omiti-los aqui para economizar espaço, mas eles devem permanecer no seu código.
    # ...
    # COPIE E COLE TODOS OS SEUS MÉTODOS DE start_new_project ATÉ set_buttons_enabled_on_process AQUI
    # ...

    def start_new_project(self):
        parent_dir = QFileDialog.getExistingDirectory(self, "Selecione a Pasta Principal para o Novo Projeto")
        if not parent_dir: return
        project_name, ok = QInputDialog.getText(self, "Novo Projeto", "Digite o nome ou número do novo projeto:")
        if ok and project_name:
            project_path = os.path.join(parent_dir, project_name)
            if os.path.exists(project_path):
                reply = QMessageBox.question(self, 'Diretório Existente', f"A pasta '{project_name}' já existe.\nDeseja usá-la como o diretório do projeto ativo?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.No: return
            else:
                try: os.makedirs(project_path)
                except OSError as e: QMessageBox.critical(self, "Erro ao Criar Pasta", f"Não foi possível criar o diretório do projeto:\n{e}"); return
            self._clear_session(clear_project_number=True)
            self.project_directory = project_path
            self.projeto_input.setText(project_name)
            self.dir_label.setText(f"Projeto Ativo: {self.project_directory}")
            self.dir_label.setStyleSheet("font-style: normal; color: #E2E8F0;") # Cor do texto do tema
            self.log_text.append(f"\n--- NOVO PROJETO INICIADO: {project_name} ---")
            self.log_text.append(f"Arquivos serão salvos em: {self.project_directory}")
            self.set_initial_button_state()

    def set_initial_button_state(self):
        is_project_active = self.project_directory is not None
        has_items = not (self.excel_df.empty and self.manual_df.empty)
        self.calculate_nesting_btn.setEnabled(is_project_active and has_items)
        self.start_project_btn.setEnabled(True)
        self.history_btn.setEnabled(True)
        self.select_file_btn.setEnabled(is_project_active)
        self.clear_excel_btn.setEnabled(is_project_active and not self.excel_df.empty)
        self.generate_code_btn.setEnabled(is_project_active)
        self.add_piece_btn.setEnabled(is_project_active)
        self.replicate_btn.setEnabled(is_project_active)
        self.add_furo_btn.setEnabled(is_project_active)
        self.process_pdf_btn.setEnabled(is_project_active and has_items)
        self.process_dxf_btn.setEnabled(is_project_active and has_items)
        self.process_all_btn.setEnabled(is_project_active and has_items)
        self.conclude_project_btn.setEnabled(is_project_active and has_items)
        self.export_excel_btn.setEnabled(is_project_active and has_items)
        self.progress_bar.setVisible(False)

    def show_history_dialog(self):
        dialog = HistoryDialog(self.history_manager, self)
        if dialog.exec_() == QDialog.Accepted:
            loaded_pieces = dialog.loaded_project_data
            if loaded_pieces:
                project_number_loaded = loaded_pieces[0].get('project_number') if loaded_pieces and 'project_number' in loaded_pieces[0] else dialog.project_list_widget.currentItem().text()
                self.start_new_project_from_history(project_number_loaded, loaded_pieces)
    
    def start_new_project_from_history(self, project_name, pieces_data):
        parent_dir = QFileDialog.getExistingDirectory(self, f"Selecione uma pasta para o projeto '{project_name}'")
        if not parent_dir: return
        project_path = os.path.join(parent_dir, project_name)
        os.makedirs(project_path, exist_ok=True)
        self._clear_session(clear_project_number=True)
        self.project_directory = project_path
        self.projeto_input.setText(project_name)
        self.excel_df = pd.DataFrame(columns=self.colunas_df)
        self.manual_df = pd.DataFrame(pieces_data)
        self.dir_label.setText(f"Projeto Ativo: {self.project_directory}"); self.dir_label.setStyleSheet("font-style: normal; color: #E2E8F0;")
        self.log_text.append(f"\n--- PROJETO DO HISTÓRICO CARREGADO: {project_name} ---")
        self.update_table_display()
        self.set_initial_button_state()

    def start_pdf_generation(self): self.start_processing(generate_pdf=True, generate_dxf=False)
    def start_dxf_generation(self): self.start_processing(generate_pdf=False, generate_dxf=True)
    def start_all_generation(self): self.start_processing(generate_pdf=True, generate_dxf=True)

    def start_processing(self, generate_pdf, generate_dxf):
        if not self.project_directory:
            QMessageBox.warning(self, "Nenhum Projeto Ativo", "Inicie um novo projeto antes de gerar arquivos."); return
        project_number = self.projeto_input.text().strip()
        if not project_number:
            QMessageBox.warning(self, "Número do Projeto Ausente", "Por favor, defina um número para o projeto ativo."); return
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        if combined_df.empty: QMessageBox.warning(self, "Aviso", "A lista de peças está vazia."); return
        self.set_buttons_enabled_on_process(False)
        self.progress_bar.setVisible(True); self.progress_bar.setValue(0); self.log_text.clear()
        self.process_thread = ProcessThread(combined_df.copy(), generate_pdf, generate_dxf, self.project_directory, project_number)
        self.process_thread.update_signal.connect(self.log_text.append)
        self.process_thread.progress_signal.connect(self.progress_bar.setValue)
        self.process_thread.finished_signal.connect(self.processing_finished)
        self.process_thread.start()

    def processing_finished(self, success, message):
        self.set_buttons_enabled_on_process(True); self.progress_bar.setVisible(False)
        msgBox = QMessageBox.information if success else QMessageBox.critical
        msgBox(self, "Concluído" if success else "Erro", message); self.statusBar().showMessage("Pronto")
    
    def conclude_project(self):
        project_number = self.projeto_input.text().strip()
        if not project_number:
            QMessageBox.warning(self, "Projeto sem Número", "O projeto ativo não tem um número definido.")
            return
        reply = QMessageBox.question(self, 'Concluir Projeto', f"Deseja salvar e concluir o projeto '{project_number}'?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
            if not combined_df.empty:
                combined_df['project_number'] = project_number
                self.history_manager.save_project(project_number, combined_df)
                self.log_text.append(f"Projeto '{project_number}' salvo no histórico.")
            self._clear_session(clear_project_number=True)
            self.project_directory = None
            self.dir_label.setText("Nenhum projeto ativo. Inicie um novo projeto."); self.dir_label.setStyleSheet("font-style: italic; color: grey;")
            self.set_initial_button_state()
            self.log_text.append(f"\n--- PROJETO '{project_number}' CONCLUÍDO ---")

    def open_nesting_dialog(self):
        if self.excel_df.empty and self.manual_df.empty:
            QMessageBox.warning(self, "Lista Vazia", "Não há peças na lista para calcular o aproveitamento.")
            return
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        rect_df = combined_df[combined_df['forma'] == 'rectangle'].copy()
        if rect_df.empty:
            QMessageBox.information(self, "Nenhuma Peça Válida", "O cálculo de aproveitamento só pode ser feito com peças da forma 'rectangle'.")
            return
        dialog = NestingDialog(rect_df, self)
        dialog.exec_()

    def export_project_to_excel(self):
        chapa_largura_str, ok1 = QInputDialog.getText(self, "Parâmetro de Aproveitamento", "Largura da Chapa (mm):", text="3000")
        if not ok1: return
        chapa_altura_str, ok2 = QInputDialog.getText(self, "Parâmetro de Aproveitamento", "Altura da Chapa (mm):", text="1500")
        if not ok2: return
        offset_str, ok3 = QInputDialog.getText(self, "Parâmetro de Aproveitamento", "Offset entre Peças (mm):", text="8")
        if not ok3: return
        try:
            chapa_largura = float(chapa_largura_str)
            chapa_altura = float(chapa_altura_str)
            offset = float(offset_str)
        except (ValueError, TypeError):
            QMessageBox.critical(self, "Erro de Entrada", "Valores de chapa e offset devem ser numéricos.")
            return
        project_number = self.projeto_input.text().strip()
        if not project_number:
            QMessageBox.warning(self, "Nenhum Projeto Ativo", "Inicie um novo projeto para poder exportá-lo.")
            return
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        if combined_df.empty:
            QMessageBox.warning(self, "Lista Vazia", "Não há peças na lista para exportar.")
            return
        default_filename = os.path.join(self.project_directory, f"CUSTO_PLASMA-LASER_V4_NOVA_{project_number}.xlsx")
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Resumo do Projeto", default_filename, "Excel Files (*.xlsx)")
        if not save_path:
            return
        self.set_buttons_enabled_on_process(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        self.log_text.append("Iniciando exportação para Excel...")
        QApplication.processEvents()
        try:
            template_path = 'CUSTO_PLASMA-LASER_V4_NOVA.xlsx'
            if not os.path.exists(template_path):
                QMessageBox.critical(self, "Template Não Encontrado", f"O arquivo modelo '{template_path}' não foi encontrado no diretório da aplicação.")
                return
            wb = load_workbook(template_path)
            ws = wb.active
            self.log_text.append("Preenchendo lista de peças...")
            QApplication.processEvents()
            start_row = 1
            while ws.cell(row=start_row, column=1).value is not None:
                start_row += 1
            total_pecas = len(combined_df)
            for index, row_data in enumerate(combined_df.iterrows()):
                row_data = row_data[1]
                current_row = start_row + index
                ws.cell(row=current_row, column=1, value=project_number)
                ws.cell(row=current_row, column=2, value=row_data.get('nome_arquivo', ''))
                ws.cell(row=current_row, column=3, value=row_data.get('qtd', 0))
                forma = str(row_data.get('forma', '')).lower()
                largura = row_data.get('largura', 0)
                altura = row_data.get('altura', 0)
                forma_map = {'circle': 'C', 'trapezoid': 'TP', 'right_triangle': 'T'}
                if forma == 'rectangle':
                    forma_abreviada = 'Q' if largura == altura and largura > 0 else 'R'
                else:
                    forma_abreviada = forma_map.get(forma, '')
                ws.cell(row=current_row, column=4, value=forma_abreviada)
                furos = row_data.get('furos', [])
                num_furos = len(furos) if isinstance(furos, list) else 0
                ws.cell(row=current_row, column=5, value=num_furos)
                diametro_furo = furos[0].get('diam', 0) if num_furos > 0 else 0
                ws.cell(row=current_row, column=6, value=diametro_furo)
                ws.cell(row=current_row, column=7, value=row_data.get('espessura', 0))
                ws.cell(row=current_row, column=8, value=largura)
                ws.cell(row=current_row, column=9, value=altura)
                self.progress_bar.setValue(int(((index + 1) / (total_pecas * 2)) * 100))
            self.log_text.append("Calculando aproveitamento de chapas...")
            QApplication.processEvents()
            rect_df = combined_df[combined_df['forma'] == 'rectangle'].copy()
            rect_df['espessura'] = rect_df['espessura'].astype(float)
            grouped = rect_df.groupby('espessura')
            current_row = 209
            ws.cell(row=current_row, column=1, value="RELATÓRIO DE APROVEITAMENTO DE CHAPA").font = wb.active['A1'].font.copy(bold=True, size=14)
            current_row += 2
            for espessura, group in grouped:
                pecas_para_calcular = []
                for _, row in group.iterrows():
                    if row['largura'] > 0 and row['altura'] > 0:
                        pecas_para_calcular.append({'largura': row['largura'] + offset, 'altura': row['altura'] + offset, 'quantidade': int(row['qtd'])})
                if not pecas_para_calcular: continue
                resultado = calcular_plano_de_corte(chapa_largura, chapa_altura, pecas_para_calcular)
                ws.cell(row=current_row, column=1, value=f"Espessura: {espessura} mm").font = wb.active['A1'].font.copy(bold=True, size=12)
                current_row += 1
                total_chapas_usadas = resultado['total_chapas']
                peso_kg = (chapa_largura/1000) * (chapa_altura/1000) * (espessura/1000) * 7.85 * 1000 * total_chapas_usadas
                ws.cell(row=current_row, column=1, value=f"Total de Chapas: {total_chapas_usadas}")
                ws.cell(row=current_row, column=2, value=f"Aproveitamento: {resultado['aproveitamento_geral']}")
                ws.cell(row=current_row, column=3, value=f"Peso Total Estimado: {peso_kg:.2f} kg").font = wb.active['A1'].font.copy(bold=True)
                current_row += 2
                for i, plano_info in enumerate(resultado['planos_unicos']):
                    ws.cell(row=current_row, column=1, value=f"Plano de Corte {i+1} (Repetir {plano_info['repeticoes']}x)").font = wb.active['A1'].font.copy(italic=True)
                    current_row += 1
                    ws.cell(row=current_row, column=2, value="Peças neste plano:")
                    current_row += 1
                    for item in plano_info['resumo_pecas']:
                        dimensoes_sem_offset = item['tipo'].split('x')
                        largura_real = float(dimensoes_sem_offset[0]) - offset
                        altura_real = float(dimensoes_sem_offset[1]) - offset
                        texto_peca = f"- {item['qtd']}x de {largura_real:.0f}x{altura_real:.0f} mm"
                        ws.cell(row=current_row, column=3, value=texto_peca)
                        current_row += 1
                    current_row += 1
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=9)
                cell = ws.cell(row=current_row, column=1)
                cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                current_row += 2
                self.progress_bar.setValue(50 + int((current_row / 400) * 50))
            self.log_text.append("Salvando arquivo Excel...")
            QApplication.processEvents()
            wb.save(save_path)
            self.progress_bar.setValue(100)
            self.log_text.append(f"Resumo do projeto salvo com sucesso em: {save_path}")
            QMessageBox.information(self, "Sucesso", f"O arquivo Excel foi salvo com sucesso em:\n{save_path}")
        except Exception as e:
            self.log_text.append(f"ERRO ao exportar para Excel: {e}")
            QMessageBox.critical(self, "Erro na Exportação", f"Ocorreu um erro ao salvar o arquivo:\n{e}")
        finally:
            self.set_buttons_enabled_on_process(True)
            self.progress_bar.setVisible(False)

    def _clear_session(self, clear_project_number=False):
        fields_to_clear = [self.nome_input, self.espessura_input, self.qtd_input, self.largura_input, self.altura_input, self.diametro_input, self.rt_base_input, self.rt_height_input, self.trapezoid_large_base_input, self.trapezoid_small_base_input, self.trapezoid_height_input, self.rep_diam_input, self.rep_offset_input, self.diametro_furo_input, self.pos_x_input, self.pos_y_input]
        if clear_project_number:
            fields_to_clear.append(self.projeto_input)
        for field in fields_to_clear:
            field.clear()
        self.furos_atuais = []
        self.update_furos_table()
        self.file_label.setText("Nenhum projeto ativo.")
        if clear_project_number: 
            self.excel_df = pd.DataFrame(columns=self.colunas_df)
            self.manual_df = pd.DataFrame(columns=self.colunas_df)
            self.update_table_display()

    def set_buttons_enabled_on_process(self, enabled):
        is_project_active = self.project_directory is not None
        has_items = not (self.excel_df.empty and self.manual_df.empty)
        self.calculate_nesting_btn.setEnabled(enabled and is_project_active and has_items)
        self.start_project_btn.setEnabled(enabled)
        self.history_btn.setEnabled(enabled)
        self.select_file_btn.setEnabled(enabled and is_project_active)
        self.clear_excel_btn.setEnabled(enabled and is_project_active and not self.excel_df.empty)
        self.generate_code_btn.setEnabled(enabled and is_project_active)
        self.add_piece_btn.setEnabled(enabled and is_project_active)
        self.replicate_btn.setEnabled(enabled and is_project_active)
        self.add_furo_btn.setEnabled(enabled and is_project_active)
        self.process_pdf_btn.setEnabled(enabled and is_project_active and has_items)
        self.process_dxf_btn.setEnabled(enabled and is_project_active and has_items)
        self.process_all_btn.setEnabled(enabled and is_project_active and has_items)
        self.conclude_project_btn.setEnabled(enabled and is_project_active and has_items)
        self.export_excel_btn.setEnabled(enabled and is_project_active and has_items)

    def update_table_display(self):
        self.set_initial_button_state()
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        
        # --- CORREÇÃO: Método robusto para limpar a tabela antes de atualizar ---
        self.pieces_table.blockSignals(True)
        while self.pieces_table.rowCount() > 0:
            self.pieces_table.removeRow(0)
        self.pieces_table.blockSignals(False)
        if combined_df.empty:
            header = self.pieces_table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(True)
            return
        self.pieces_table.setRowCount(len(combined_df))
        self.pieces_table.verticalHeader().setDefaultSectionSize(40) # <<< ALTURA DE LINHA OTIMIZADA >>>
        for i, row in combined_df.iterrows():
            for j, col in enumerate(self.colunas_df):
                value = row.get(col)
                if col == 'furos' and isinstance(value, list):
                    display_value = f"{len(value)} Furo(s)"
                elif pd.isna(value) or value == 0:
                    display_value = '-'
                else:
                    display_value = str(value)
                item = QTableWidgetItem(display_value)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft) # <<< ALINHAMENTO MELHORADO >>>
                self.pieces_table.setItem(i, j, item)

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(5, 0, 5, 0)
            action_layout.setSpacing(5)
            edit_btn, delete_btn = QPushButton("Editar"), QPushButton("Excluir")
            edit_btn.clicked.connect(lambda _, r=i: self.edit_row(r))
            delete_btn.clicked.connect(lambda _, r=i: self.delete_row(r))
            action_layout.addWidget(edit_btn)
            action_layout.addWidget(delete_btn)
            self.pieces_table.setCellWidget(i, len(self.colunas_df), action_widget)

        header = self.pieces_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
    
    def edit_row(self, row_index):
        len_excel = len(self.excel_df)
        is_from_excel = row_index < len_excel
        df_source = self.excel_df if is_from_excel else self.manual_df
        local_index = row_index if is_from_excel else row_index - len_excel
        piece_data = df_source.iloc[local_index]
        self.nome_input.setText(str(piece_data.get('nome_arquivo', '')))
        self.espessura_input.setText(str(piece_data.get('espessura', '')))
        self.qtd_input.setText(str(piece_data.get('qtd', '')))
        shape = piece_data.get('forma', '')
        index = self.forma_combo.findText(shape, Qt.MatchFixedString)
        if index >= 0: self.forma_combo.setCurrentIndex(index)
        self.largura_input.setText(str(piece_data.get('largura', '')))
        self.altura_input.setText(str(piece_data.get('altura', '')))
        self.diametro_input.setText(str(piece_data.get('diametro', '')))
        self.rt_base_input.setText(str(piece_data.get('rt_base', '')))
        self.rt_height_input.setText(str(piece_data.get('rt_height', '')))
        self.trapezoid_large_base_input.setText(str(piece_data.get('trapezoid_large_base', '')))
        self.trapezoid_small_base_input.setText(str(piece_data.get('trapezoid_small_base', '')))
        self.trapezoid_height_input.setText(str(piece_data.get('trapezoid_height', '')))
        self.furos_atuais = piece_data.get('furos', []).copy() if isinstance(piece_data.get('furos'), list) else []
        self.update_furos_table()
        df_source.drop(df_source.index[local_index], inplace=True)
        df_source.reset_index(drop=True, inplace=True)
        self.log_text.append(f"Peça '{piece_data['nome_arquivo']}' carregada para edição.")
        self.update_table_display()
    
    def delete_row(self, row_index):
        len_excel = len(self.excel_df)
        is_from_excel = row_index < len_excel
        df_source = self.excel_df if is_from_excel else self.manual_df
        local_index = row_index if is_from_excel else row_index - len_excel
        piece_name = df_source.iloc[local_index]['nome_arquivo']
        df_source.drop(df_source.index[local_index], inplace=True)
        df_source.reset_index(drop=True, inplace=True)
        self.log_text.append(f"Peça '{piece_name}' removida.")
        self.update_table_display()
    
    def generate_piece_code(self):
        project_number = self.projeto_input.text().strip()
        if not project_number: QMessageBox.warning(self, "Campo Obrigatório", "Inicie um projeto para definir o 'Nº do Projeto'."); return
        new_code = self.code_generator.generate_new_code(project_number, prefix='DES')
        if new_code: self.nome_input.setText(new_code); self.log_text.append(f"Código '{new_code}' gerado para o projeto '{project_number}'.")
    
    def add_manual_piece(self):
        try:
            nome = self.nome_input.text().strip()
            if not nome: QMessageBox.warning(self, "Campo Obrigatório", "'Nome/ID da Peça' é obrigatório."); return
            new_piece = {'furos': self.furos_atuais.copy()}
            for col in self.colunas_df:
                if col != 'furos': new_piece[col] = 0.0
            new_piece.update({'nome_arquivo': nome, 'forma': self.forma_combo.currentText()})
            fields_map = { 'espessura': self.espessura_input, 'qtd': self.qtd_input, 'largura': self.largura_input, 'altura': self.altura_input, 'diametro': self.diametro_input, 'rt_base': self.rt_base_input, 'rt_height': self.rt_height_input, 'trapezoid_large_base': self.trapezoid_large_base_input, 'trapezoid_small_base': self.trapezoid_small_base_input, 'trapezoid_height': self.trapezoid_height_input }
            for key, field in fields_map.items():
                new_piece[key] = float(field.text().replace(',', '.')) if field.text() else 0.0
            self.manual_df = pd.concat([self.manual_df, pd.DataFrame([new_piece])], ignore_index=True)
            self.log_text.append(f"Peça '{nome}' adicionada/atualizada.")
            self._clear_session(clear_project_number=False)
            self.update_table_display()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Campos numéricos devem conter números válidos.")
    
    def select_file(self):
        if not self.project_directory:
            QMessageBox.warning(self, "Nenhum Projeto Ativo", "Inicie um projeto antes de carregar uma planilha.")
            return
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                df = pd.read_excel(file_path, header=0, decimal=','); df.columns = df.columns.str.strip().str.lower()
                df = df.loc[:, ~df.columns.duplicated()]
                for col in self.colunas_df:
                    if col not in df.columns: df[col] = pd.NA
                if 'furos' in df.columns:
                    def parse_furos(x):
                        if isinstance(x, list): return x
                        if isinstance(x, str) and x.startswith('['):
                            try: return json.loads(x.replace("'", "\""))
                            except json.JSONDecodeError: return []
                        return []
                    df['furos'] = df['furos'].apply(parse_furos)
                else:
                    df['furos'] = [[] for _ in range(len(df))]
                numeric_cols = [col for col in self.colunas_df if col != 'furos' and col != 'forma' and col != 'nome_arquivo']
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                self.excel_df = df[self.colunas_df]
                self.file_label.setText(f"Planilha: {os.path.basename(file_path)}"); self.update_table_display()
            except Exception as e: QMessageBox.critical(self, "Erro de Leitura", f"Falha ao ler o arquivo: {e}")
    
    def clear_excel_data(self):
        self.excel_df = pd.DataFrame(columns=self.colunas_df); self.file_label.setText("Nenhuma planilha selecionada"); self.update_table_display()
    
    def replicate_holes(self):
        try:
            if self.forma_combo.currentText() != 'rectangle': QMessageBox.warning(self, "Função Indisponível", "Replicação disponível apenas para Retângulos."); return
            largura, altura = float(self.largura_input.text().replace(',', '.')), float(self.altura_input.text().replace(',', '.'))
            diam, offset = float(self.rep_diam_input.text().replace(',', '.')), float(self.rep_offset_input.text().replace(',', '.'))
            if (offset * 2) >= largura or (offset * 2) >= altura: QMessageBox.warning(self, "Offset Inválido", "Offset excede as dimensões da peça."); return
            furos = [{'diam': diam, 'x': offset, 'y': offset}, {'diam': diam, 'x': largura - offset, 'y': offset}, {'diam': diam, 'x': largura - offset, 'y': altura - offset}, {'diam': diam, 'x': offset, 'y': altura - offset}]
            self.furos_atuais.extend(furos); self.update_furos_table()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Largura, Altura, Diâmetro e Offset devem ser números válidos.")
    
    def update_dimension_fields(self, shape):
        shape = shape.lower()
        is_rect, is_circ, is_tri, is_trap = shape == 'rectangle', shape == 'circle', shape == 'right_triangle', shape == 'trapezoid'
        for w in self.largura_row + self.altura_row: w.setVisible(is_rect)
        for w in self.diametro_row: w.setVisible(is_circ)
        for w in self.rt_base_row + self.rt_height_row: w.setVisible(is_tri)
        for w in self.trap_large_base_row + self.trap_small_base_row + self.trap_height_row: w.setVisible(is_trap)
        self.rep_group.setEnabled(is_rect)
    
    def add_furo_temp(self):
        try:
            diam, pos_x, pos_y = float(self.diametro_furo_input.text().replace(',', '.')), float(self.pos_x_input.text().replace(',', '.')), float(self.pos_y_input.text().replace(',', '.'))
            if diam <= 0: QMessageBox.warning(self, "Valor Inválido", "Diâmetro do furo deve ser maior que zero."); return
            self.furos_atuais.append({'diam': diam, 'x': pos_x, 'y': pos_y}); self.update_furos_table()
            for field in [self.diametro_furo_input, self.pos_x_input, self.pos_y_input]: field.clear()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Campos de furo devem ser números válidos.")
    
    def update_furos_table(self):
        self.furos_table.setRowCount(0); self.furos_table.setRowCount(len(self.furos_atuais))
        for i, furo in enumerate(self.furos_atuais):
            self.furos_table.setItem(i, 0, QTableWidgetItem(str(furo['diam'])))
            self.furos_table.setItem(i, 1, QTableWidgetItem(str(furo['x'])))
            self.furos_table.setItem(i, 2, QTableWidgetItem(str(furo['y'])))
            delete_btn = QPushButton("Excluir")
            delete_btn.clicked.connect(lambda _, r=i: self.delete_furo_temp(r))
            self.furos_table.setCellWidget(i, 3, delete_btn)
        self.furos_table.resizeColumnsToContents()
    
    def delete_furo_temp(self, row_index):
        del self.furos_atuais[row_index]
        self.update_furos_table()

# =============================================================================
# PONTO DE ENTRADA DA APLICAÇÃO
# =============================================================================
def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(INOVA_PROCESS_STYLE) # Aplica o tema escuro em toda a aplicação
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()