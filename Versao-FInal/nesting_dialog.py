import sys
import os
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QFormLayout, QLineEdit, 
                             QPushButton, QDialogButtonBox, QMessageBox, 
                             QGroupBox, QLabel, QWidget, QHBoxLayout, QScrollArea,
                             QFileDialog)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QColor, QPen, QBrush
# Importações para gerar PDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pdf_generator
# Importe sua função de cálculo
from calculo_cortes import calcular_plano_de_corte 

# Classe para desenhar a visualização do plano de corte
class CuttingPlanWidget(QWidget):
    def __init__(self, chapa_largura, chapa_altura, plano, parent=None):
        super().__init__(parent)
        self.chapa_largura = chapa_largura
        self.chapa_altura = chapa_altura
        self.plano = plano
        self.setMinimumSize(400, 600)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Escala para caber no widget
        scale_w = self.width() / self.chapa_largura
        scale_h = self.height() / self.chapa_altura
        scale = min(scale_w, scale_h) * 0.95 # 5% de margem

        # Centraliza o desenho
        offset_x = (self.width() - (self.chapa_largura * scale)) / 2
        offset_y = (self.height() - (self.chapa_altura * scale)) / 2

        # 1. Desenha a chapa
        painter.setPen(QPen(QColor("#CCCCCC")))
        painter.setBrush(QBrush(QColor("#FFFFFF")))
        painter.drawRect(int(offset_x), int(offset_y), int(self.chapa_largura * scale), int(self.chapa_altura * scale))

        # 2. Desenha as peças
        for peca in self.plano:
            x, y, w, h = peca['x'], peca['y'], peca['largura'], peca['altura']
            
            # Posições e dimensões escaladas
            rect_x = int(offset_x + x * scale)
            rect_y = int(offset_y + y * scale)
            rect_w = int(w * scale)
            rect_h = int(h * scale)

            painter.setPen(QPen(QColor("#333333")))
            painter.setBrush(QBrush(QColor("#A94442"))) # Um tom de vermelho
            painter.drawRect(rect_x, rect_y, rect_w, rect_h)


class PlanVisualizationDialog(QDialog):
    def __init__(self, chapa_largura, chapa_altura, plano, parent=None):
        super().__init__(parent)
        self.chapa_largura = chapa_largura
        self.chapa_altura = chapa_altura
        self.plano = plano
        self.setWindowTitle("Visualização Detalhada do Plano de Corte")
        self.setMinimumSize(600, 800)
        
        layout = QVBoxLayout(self)
        cutting_widget = CuttingPlanWidget(chapa_largura, chapa_altura, plano)
        layout.addWidget(cutting_widget)

        # Botões
        buttons_layout = QHBoxLayout()
        btn_export_pdf = QPushButton("Exportar para PDF")
        btn_export_pdf.clicked.connect(self.export_to_pdf)
        
        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(self.accept)

        buttons_layout.addWidget(btn_export_pdf)
        buttons_layout.addStretch()
        buttons_layout.addWidget(btn_close)
        layout.addLayout(buttons_layout)

    def export_to_pdf(self):
        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "Plano_de_Corte.pdf")
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF do Plano de Corte", default_path, "PDF Files (*.pdf)")
        if save_path:
            c = canvas.Canvas(save_path, pagesize=A4)
            pdf_generator.gerar_pdf_plano_de_corte(c, self.chapa_largura, self.chapa_altura, self.plano)
            c.save()
            QMessageBox.information(self, "Sucesso", f"PDF do plano de corte salvo em:\n{save_path}")


class NestingDialog(QDialog):
    def __init__(self, dataframe, parent=None):
        super().__init__(parent)
        self.df = dataframe
        self.calculation_results = None # Armazena os resultados completos
        self.setWindowTitle("Cálculo de Aproveitamento de Chapa")
        self.setMinimumWidth(600)

        # Layout principal
        self.main_layout = QVBoxLayout(self)

        # --- Grupo de Inputs do Usuário ---
        input_group = QGroupBox("Parâmetros")
        form_layout = QFormLayout()
        self.chapa_largura_input = QLineEdit("3000")
        self.chapa_altura_input = QLineEdit("1500")
        self.offset_input = QLineEdit("8")
        form_layout.addRow("Largura da Chapa (mm):", self.chapa_largura_input)
        form_layout.addRow("Altura da Chapa (mm):", self.chapa_altura_input)
        form_layout.addRow("Offset entre Peças (mm):", self.offset_input)
        input_group.setLayout(form_layout)
        self.main_layout.addWidget(input_group)
        
        # Botões de Ação
        action_layout = QHBoxLayout()
        self.calculate_btn = QPushButton("Calcular")
        self.calculate_btn.clicked.connect(self.run_calculation)
        self.export_report_btn = QPushButton("Exportar Relatório PDF")
        self.export_report_btn.clicked.connect(self.export_full_report_to_pdf)
        self.export_report_btn.setEnabled(False) # Desabilitado até o cálculo ser feito
        action_layout.addWidget(self.calculate_btn)
        action_layout.addWidget(self.export_report_btn)
        self.main_layout.addLayout(action_layout)

        # --- Área de Resultados ---
        results_group = QGroupBox("Resultados")
        self.results_layout = QVBoxLayout()
        results_group.setLayout(self.results_layout)
        
        # Adiciona uma área de rolagem para os resultados
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self.results_scroll_layout = QVBoxLayout(scroll_content)
        scroll.setWidget(scroll_content)

        self.main_layout.addWidget(results_group)
        self.main_layout.addWidget(scroll)

    def run_calculation(self):
        try:
            chapa_largura = float(self.chapa_largura_input.text())
            chapa_altura = float(self.chapa_altura_input.text())
            offset = float(self.offset_input.text())
        except ValueError:
            QMessageBox.critical(self, "Erro de Entrada", "Por favor, insira valores numéricos válidos.")
            return

        # Limpa resultados anteriores
        for i in reversed(range(self.results_scroll_layout.count())): 
            self.results_scroll_layout.itemAt(i).widget().setParent(None)
        
        self.calculation_results = {} # Limpa e prepara para novos resultados
        self.export_report_btn.setEnabled(False)

        # Filtra apenas peças retangulares e agrupa por espessura
        rect_df = self.df[self.df['forma'] == 'rectangle'].copy()
        rect_df['espessura'] = rect_df['espessura'].astype(float)

        grouped = rect_df.groupby('espessura')

        if len(grouped) == 0:
            self.results_scroll_layout.addWidget(QLabel("Nenhuma peça retangular encontrada para o cálculo."))
            return

        for espessura, group in grouped:
            pecas_para_calcular = []
            for _, row in group.iterrows():
                # Valida se largura e altura são maiores que zero
                if row['largura'] > 0 and row['altura'] > 0:
                    pecas_para_calcular.append({
                        # Adiciona o offset às dimensões da peça
                        'largura': row['largura'] + offset,
                        'altura': row['altura'] + offset,
                        'quantidade': int(row['qtd'])
                    })
            
            if not pecas_para_calcular:
                continue

            try:
                # Chama a função de cálculo importada
                resultado = calcular_plano_de_corte(chapa_largura, chapa_altura, pecas_para_calcular)
                self.calculation_results[espessura] = resultado # Armazena o resultado
                self.display_results_for_thickness(espessura, resultado, chapa_largura, chapa_altura)
            except Exception as e:
                QMessageBox.critical(self, f"Erro no Cálculo (Espessura {espessura}mm)", str(e))
        
        # Habilita o botão de exportar se houver resultados
        if self.calculation_results:
            self.export_report_btn.setEnabled(True)

    def export_full_report_to_pdf(self):
        if not self.calculation_results:
            QMessageBox.warning(self, "Sem Dados", "Nenhum resultado de cálculo para exportar. Por favor, clique em 'Calcular' primeiro.")
            return

        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "Relatorio_Aproveitamento.pdf")
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório de Aproveitamento", default_path, "PDF Files (*.pdf)")
        if save_path:
            c = canvas.Canvas(save_path, pagesize=A4)
            pdf_generator.gerar_pdf_aproveitamento_completo(c, self.calculation_results, float(self.chapa_largura_input.text()), float(self.chapa_altura_input.text()))
            c.save()
            QMessageBox.information(self, "Sucesso", f"Relatório PDF salvo em:\n{save_path}")

    def display_results_for_thickness(self, espessura, resultado, chapa_w, chapa_h):
        # Cria um grupo para cada espessura
        group_box = QGroupBox(f"Espessura: {espessura} mm")
        group_layout = QVBoxLayout()

        # Adiciona informações gerais
        info_label = QLabel(f"Total de Chapas: {resultado['total_chapas']} | Aproveitamento Geral: {resultado['aproveitamento_geral']}")
        info_label.setStyleSheet("font-weight: bold;")
        group_layout.addWidget(info_label)

        # Lista os planos de corte únicos
        for i, plano_info in enumerate(resultado['planos_unicos']):
            plano_layout = QHBoxLayout()
            
            resumo_pecas_str = ", ".join([f"{p['qtd']}x ({p['tipo']})" for p in plano_info['resumo_pecas']])
            
            plan_label = QLabel(f"Plano {i+1}: {plano_info['repeticoes']}x | Peças: {resumo_pecas_str}")
            
            view_btn = QPushButton("Ver Detalhes")
            # Usa lambda para capturar os dados corretos do plano para o dialog
            view_btn.clicked.connect(lambda _, p=plano_info['plano'], w=chapa_w, h=chapa_h: self.show_plan_visualization(p, w, h))
            
            plano_layout.addWidget(plan_label)
            plano_layout.addStretch()
            plano_layout.addWidget(view_btn)
            
            group_layout.addLayout(plano_layout)

        group_box.setLayout(group_layout)
        self.results_scroll_layout.addWidget(group_box)

    def show_plan_visualization(self, plano, chapa_w, chapa_h):
        dialog = PlanVisualizationDialog(chapa_w, chapa_h, plano, self)
        dialog.exec_()