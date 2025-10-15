import sys
import os
import colorsys
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

# --- INÍCIO: FUNÇÃO PARA GERAR CORES DISTINTAS ---
def generate_distinct_colors(n):
    """Gera N cores visualmente distintas."""
    colors = []
    for i in range(n):
        hue = i / n
        # Usamos saturação e valor altos para cores vibrantes, mas não totalmente saturadas para não cansar a vista.
        saturation = 0.85
        value = 0.9
        rgb = colorsys.hsv_to_rgb(hue, saturation, value)
        colors.append(QColor(int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255)))
    return colors
# --- FIM: FUNÇÃO PARA GERAR CORES DISTINTAS ---

# Classe para desenhar a visualização do plano de corte
class CuttingPlanWidget(QWidget):
    def __init__(self, chapa_largura, chapa_altura, plano, color_map, parent=None):
        super().__init__(parent)
        self.chapa_largura, self.chapa_altura = chapa_largura, chapa_altura
        self.plano = plano
        self.color_map = color_map
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
        painter.setPen(QPen(QColor("#333333")))
        for peca in self.plano:
            x, y, w, h, tipo_key = peca['x'], peca['y'], peca['largura'], peca['altura'], peca['tipo_key']
            
            rect_x = int(offset_x + x * scale)
            rect_y = int(offset_y + y * scale)
            rect_w = int(w * scale)
            rect_h = int(h * scale)

            # Define a cor da peça
            painter.setBrush(QBrush(self.color_map.get(tipo_key, QColor("#A94442"))))
            painter.drawRect(rect_x, rect_y, rect_w, rect_h)

            # 3. Desenha os furos dentro da peça
            furos = peca.get('furos', [])
            if furos:
                painter.setBrush(QBrush(QColor("#FFFFFF"))) # Furos brancos
                for furo in furos:
                    furo_x = rect_x + (furo['x'] * scale)
                    furo_y = rect_y + (furo['y'] * scale)
                    furo_diam = furo['diam'] * scale
                    # Desenha o círculo do furo centralizado na sua coordenada
                    painter.drawEllipse(int(furo_x - furo_diam / 2), int(furo_y - furo_diam / 2), int(furo_diam), int(furo_diam))


class PlanVisualizationDialog(QDialog):
    def __init__(self, chapa_largura, chapa_altura, plano_info, offset, color_map, parent=None):
        super().__init__(parent)
        self.chapa_largura = chapa_largura
        self.chapa_altura = chapa_altura
        self.plano = plano_info['plano']
        self.resumo_pecas = plano_info['resumo_pecas']
        self.offset = offset
        self.color_map = color_map
        self.setWindowTitle("Visualização Detalhada do Plano de Corte")
        self.setMinimumSize(600, 800)
        
        layout = QVBoxLayout(self)

        # --- INÍCIO: NOVOS LABELS DE INFORMAÇÃO ---
        info_group = QGroupBox("Detalhes do Plano")
        info_layout = QVBoxLayout()

        chapa_label = QLabel(f"<b>Dimensões da Chapa:</b> {self.chapa_largura} x {self.chapa_altura} mm")
        info_layout.addWidget(chapa_label)

        pecas_label_titulo = QLabel("<b>Peças neste plano:</b>")
        info_layout.addWidget(pecas_label_titulo)

        for item in self.resumo_pecas:
            dimensoes_com_offset = item['tipo'].split('x')
            largura_real = float(dimensoes_com_offset[0]) - self.offset
            altura_real = float(dimensoes_com_offset[1]) - self.offset
            texto_peca = f"- {item['qtd']}x de {largura_real:.0f} x {altura_real:.0f} mm"
            info_layout.addWidget(QLabel(texto_peca))

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        # --- FIM: NOVOS LABELS DE INFORMAÇÃO ---

        cutting_widget = CuttingPlanWidget(chapa_largura, chapa_altura, self.plano, color_map)
        layout.addWidget(cutting_widget)

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
            pdf_generator.gerar_pdf_plano_de_corte(c, self.chapa_largura, self.chapa_altura, self.plano, self.color_map)
            c.save()
            QMessageBox.information(self, "Sucesso", f"PDF do plano de corte salvo em:\n{save_path}")


class NestingDialog(QDialog):
    def __init__(self, dataframe, parent=None):
        super().__init__(parent)
        self.df = dataframe
        self.calculation_results = None # Armazena os resultados completos
        self.color_map = {} # Armazena o mapa de cores por tipo de peça
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
        # --- INÍCIO: BOTÃO DE EXPANSÃO ---
        title_bar_layout = QHBoxLayout()
        title_bar_layout.addStretch()
        self.toggle_fullscreen_btn = QPushButton("🗖") # Símbolo de maximizar/restaurar
        self.toggle_fullscreen_btn.setFixedSize(30, 30)
        self.toggle_fullscreen_btn.setToolTip("Maximizar / Restaurar Janela")
        self.toggle_fullscreen_btn.clicked.connect(self.toggle_fullscreen)
        title_bar_layout.addWidget(self.toggle_fullscreen_btn)
        self.main_layout.addLayout(title_bar_layout)
        # --- FIM: BOTÃO DE EXPANSÃO ---

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

    def toggle_fullscreen(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()


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
                        'quantidade': int(row['qtd']),
                        'furos': row.get('furos', []) # Passa a informação dos furos
                    })
            
            # --- INÍCIO: GERAÇÃO DO MAPA DE CORES ---
            tipos_de_peca_unicos = group.apply(lambda r: f"{r['largura'] + offset}x{r['altura'] + offset}", axis=1).unique()
            cores = generate_distinct_colors(len(tipos_de_peca_unicos))
            self.color_map = {tipo: cor for tipo, cor in zip(tipos_de_peca_unicos, cores)}
            # --- FIM: GERAÇÃO DO MAPA DE CORES ---

            if not pecas_para_calcular:
                continue

            try:
                # Chama a função de cálculo importada
                resultado = calcular_plano_de_corte(chapa_largura, chapa_altura, pecas_para_calcular) # Furos são passados aqui
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
            # --- MUDANÇA: Passa o dicionário 'plano_info' completo e o offset ---
            view_btn.clicked.connect(lambda _, p_info=plano_info, w=chapa_w, h=chapa_h: self.show_plan_visualization(p_info, w, h, self.color_map))
            
            plano_layout.addWidget(plan_label)
            plano_layout.addStretch()
            plano_layout.addWidget(view_btn)
            
            group_layout.addLayout(plano_layout)

        group_box.setLayout(group_layout)
        self.results_scroll_layout.addWidget(group_box)

    def show_plan_visualization(self, plano_info, chapa_w, chapa_h, color_map):
        offset = float(self.offset_input.text())
        dialog = PlanVisualizationDialog(chapa_w, chapa_h, plano_info, offset, color_map, self)
        dialog.exec_()