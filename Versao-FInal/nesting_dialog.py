import sys
import os
import colorsys
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QFormLayout, QLineEdit, 
                             QPushButton, QDialogButtonBox, QMessageBox, 
                             QGroupBox, QLabel, QWidget, QHBoxLayout, QScrollArea,
                             QFileDialog)
import logging
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPointF
from PyQt5.QtGui import QPainter, QColor, QPen, QBrush, QPainterPath
# Importações para gerar PDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pdf_generator
# Importe sua função de cálculo
from calculo_cortes import calcular_plano_de_corte, status_signaler

# --- INÍCIO: CLASSE DA THREAD DE CÁLCULO ---
class CalculationThread(QThread):
    """Thread para executar o cálculo de nesting em segundo plano."""
    # Sinal que emite o resultado para uma espessura: (espessura, resultado_dict)
    result_ready = pyqtSignal(float, dict)
    # Sinal emitido quando todos os cálculos terminam
    finished = pyqtSignal()
    # Sinal para reportar erros: (titulo_erro, mensagem_erro)
    error = pyqtSignal(str, str)
    # Sinal para atualizações de status em tempo real
    status_update = pyqtSignal(str)

    def __init__(self, chapa_largura, chapa_altura, offset, grouped_df, parent=None):
        super().__init__(parent)
        self.chapa_largura = chapa_largura
        self.chapa_altura = chapa_altura
        self.offset = offset
        self.grouped_df = grouped_df

    def run(self):
        try:
            logging.info("Thread de cálculo iniciada.")
            for espessura, group in self.grouped_df:
                pecas_para_calcular = []
                # --- INÍCIO: LÓGICA PARA INCLUIR CÍRCULOS NO CÁLCULO ---
                for _, row in group.iterrows():
                    if row['forma'] == 'rectangle' and row['largura'] > 0 and row['altura'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'rectangle',
                            'largura': row['largura'] + self.offset,
                            'altura': row['altura'] + self.offset,
                            'quantidade': int(row['qtd']),
                            'furos': row.get('furos', [])
                        })
                    elif row['forma'] == 'circle' and row['diametro'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'circle',
                            'largura': row['diametro'] + self.offset, # Bounding box
                            'altura': row['diametro'] + self.offset, # Bounding box
                            'diametro': row['diametro'], # Diâmetro original
                            'quantidade': int(row['qtd']),
                            'furos': row.get('furos', [])
                        })
                    elif row['forma'] == 'right_triangle' and row['rt_base'] > 0 and row['rt_height'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'right_triangle',
                            'largura': row['rt_base'] + self.offset, # Bounding box
                            'altura': row['rt_height'] + self.offset, # Bounding box
                            'quantidade': int(row['qtd']),
                            'furos': [] # Furos em triângulos não implementado
                        })
                    elif row['forma'] == 'trapezoid' and row['trapezoid_large_base'] > 0 and row['trapezoid_height'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'trapezoid',
                            'largura': row['trapezoid_large_base'] + self.offset, # Bounding box
                            'altura': row['trapezoid_height'] + self.offset, # Bounding box
                            'small_base': row['trapezoid_small_base'] + self.offset,
                            'quantidade': int(row['qtd']),
                            'furos': row.get('furos', [])
                        })
                # --- FIM: LÓGICA PARA INCLUIR CÍRCULOS ---

                if not pecas_para_calcular:
                    continue
                
                logging.debug(f"Iniciando cálculo para espessura {espessura} com {len(pecas_para_calcular)} tipos de peças.")
                # Chama a função de cálculo pesada
                resultado = calcular_plano_de_corte(self.chapa_largura, self.chapa_altura, pecas_para_calcular, self.status_update)
                logging.debug(f"Cálculo para espessura {espessura} concluído. Emitindo resultado.")
                self.result_ready.emit(espessura, resultado)
        except Exception as e:
            logging.error(f"Erro na thread de cálculo: {e}", exc_info=True) # exc_info=True para logar o traceback
            self.error.emit(f"Erro no Cálculo (Espessura {espessura}mm)", str(e))
        finally:
            logging.info("Thread de cálculo finalizada.")
            self.finished.emit()

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

            # --- INÍCIO: DESENHO CONDICIONAL (RETÂNGULO OU CÍRCULO) ---
            forma = peca.get('forma', 'rectangle')
            cor_peca = self.color_map.get(tipo_key, QColor("#A94442"))
            painter.setBrush(QBrush(cor_peca))

            if forma == 'circle':
                # Para círculos, desenha o círculo original dentro do seu bounding box alocado
                diametro_original = peca.get('diametro', 0)
                raio_desenhado = (diametro_original * scale) / 2
                centro_x = rect_x + rect_w / 2
                centro_y = rect_y + rect_h / 2
                painter.drawEllipse(int(centro_x - raio_desenhado), int(centro_y - raio_desenhado), int(raio_desenhado * 2), int(raio_desenhado * 2))
            elif forma == 'paired_triangle':
                # Desenha dois triângulos opostos para representar o par
                path1 = QPainterPath()
                path1.moveTo(rect_x, rect_y)
                path1.lineTo(rect_x + rect_w, rect_y)
                path1.lineTo(rect_x, rect_y + rect_h)
                path1.closeSubpath()
                path2 = QPainterPath()
                path2.moveTo(rect_x + rect_w, rect_y + rect_h)
                path2.lineTo(rect_x, rect_y + rect_h)
                path2.lineTo(rect_x + rect_w, rect_y)
                path2.closeSubpath()
                painter.drawPath(path1); painter.drawPath(path2)
            elif forma == 'paired_trapezoid':
                orig_dims = peca.get('orig_dims')
                if orig_dims:
                    large_base_scaled = orig_dims['large_base'] * scale
                    small_base_scaled = orig_dims['small_base'] * scale
                    height_scaled = orig_dims['height'] * scale
                    offset_x_trap = (large_base_scaled - small_base_scaled) / 2

                    # Trapézio 1 (normal)
                    path1 = QPainterPath()
                    path1.moveTo(rect_x, rect_y); path1.lineTo(rect_x + large_base_scaled, rect_y); path1.lineTo(rect_x + large_base_scaled - offset_x_trap, rect_y + height_scaled); path1.lineTo(rect_x + offset_x_trap, rect_y + height_scaled); path1.closeSubpath()
                    
                    # Trapézio 2 (rotacionado 180 graus e deslocado)
                    # --- CORREÇÃO DA LÓGICA DE DESENHO DO SEGUNDO TRAPÉZIO ---
                    x_base2 = rect_x + large_base_scaled
                    path2 = QPainterPath()
                    # Ponto inferior esquerdo do 2º trapézio (coincide com inferior direito do 1º)
                    path2.moveTo(rect_x + large_base_scaled, rect_y)
                    # Ponto inferior direito do 2º trapézio
                    path2.lineTo(rect_x + large_base_scaled + small_base_scaled, rect_y)
                    # Ponto superior direito do 2º trapézio
                    path2.lineTo(rect_x + large_base_scaled + offset_x_trap, rect_y + height_scaled)
                    # Ponto superior esquerdo do 2º trapézio (coincide com superior direito do 1º)
                    path2.lineTo(rect_x + large_base_scaled - offset_x_trap, rect_y + height_scaled)
                    path2.closeSubpath()
                    painter.drawPath(path1); painter.drawPath(path2)
            else: # 'rectangle'
                painter.drawRect(rect_x, rect_y, rect_w, rect_h)
            # --- FIM: DESENHO CONDICIONAL ---

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

        # 4. Desenha as sobras
        sobras = self.parent().plano_sobras # Pega as sobras do diálogo pai
        if sobras:
            font = painter.font()
            font.setPointSize(7)
            painter.setFont(font)

            for sobra in sobras:
                if sobra.get('tipo_sobra') == 'aproveitavel':
                    painter.setBrush(QBrush(QColor(100, 100, 100, 150))) # Cinza escuro semi-transparente
                    painter.setPen(QPen(QColor("#333333"), 1, Qt.DashLine))
                else:
                    painter.setBrush(QBrush(QColor(230, 230, 230, 120))) # Cinza claro semi-transparente
                    painter.setPen(QPen(QColor("#888888"), 1, Qt.DashLine))
                rect_x, rect_y = int(offset_x + sobra['x'] * scale), int(offset_y + sobra['y'] * scale)
                rect_w, rect_h = int(sobra['largura'] * scale), int(sobra['altura'] * scale)
                painter.drawRect(rect_x, rect_y, rect_w, rect_h)
                painter.drawText(rect_x + 4, rect_y + 12, f"{sobra['largura']:.0f}x{sobra['altura']:.0f}")

class PlanVisualizationDialog(QDialog):
    def __init__(self, chapa_largura, chapa_altura, plano_info, offset, color_map, parent=None):
        super().__init__(parent)
        self.chapa_largura = chapa_largura
        self.chapa_altura = chapa_altura
        self.plano = plano_info['plano']
        self.repeticoes = plano_info['repeticoes']
        self.plano_sobras = plano_info.get('sobras', []) # Armazena as sobras
        self.resumo_pecas = plano_info['resumo_pecas']
        self.offset = offset
        self.color_map = color_map
        self.setWindowTitle("Visualização Detalhada do Plano de Corte")
        self.setMinimumSize(600, 800)
        
        layout = QVBoxLayout(self)

        # --- INÍCIO: CONTAINER PARA DETALHES ---
        # Agrupa todos os QGroupBox de detalhes em um único widget para facilitar o toggle
        self.details_container = QWidget()
        details_layout = QVBoxLayout(self.details_container)
        details_layout.setContentsMargins(0, 0, 0, 0)
        # --- INÍCIO: NOVOS LABELS DE INFORMAÇÃO ---
        info_group = QGroupBox("Detalhes do Plano")
        info_layout = QVBoxLayout()

        chapa_label = QLabel(f"<b>Dimensões da Chapa:</b> {self.chapa_largura} x {self.chapa_altura} mm")
        info_layout.addWidget(chapa_label)

        repeticoes_label = QLabel(f"<b>Repetições deste Plano:</b> {self.repeticoes}x")
        info_layout.addWidget(repeticoes_label)

        pecas_label_titulo = QLabel("<b>Peças neste plano:</b>")
        info_layout.addWidget(pecas_label_titulo)

        for item in self.resumo_pecas:
            # --- CORREÇÃO: Exibe a identificação da peça diretamente ---
            texto_peca = f"- {item['qtd']}x de {item['tipo']}"
            info_layout.addWidget(QLabel(texto_peca))
        info_group.setLayout(info_layout)
        details_layout.addWidget(info_group)

        # --- INÍCIO: SEPARAÇÃO DAS SOBRAS NA VISUALIZAÇÃO ---
        if self.plano_sobras:
            sobras_aproveitaveis = [s for s in self.plano_sobras if s.get('tipo_sobra') == 'aproveitavel']
            sobras_sucata = [s for s in self.plano_sobras if s.get('tipo_sobra') != 'aproveitavel']

            if sobras_aproveitaveis:
                aprov_group = QGroupBox("Sobras Aproveitáveis")
                aprov_layout = QVBoxLayout()
                for i, sobra in enumerate(sobras_aproveitaveis):
                    aprov_layout.addWidget(QLabel(f"- Retalho {i+1}: {sobra['largura']:.0f} x {sobra['altura']:.0f} mm"))
                aprov_group.setLayout(aprov_layout)
                details_layout.addWidget(aprov_group)
            
            if sobras_sucata:
                sucata_group = QGroupBox("Sobras (Sucata)")
                sucata_layout = QVBoxLayout()
                for i, sobra in enumerate(sobras_sucata):
                    sucata_layout.addWidget(QLabel(f"- Retalho {i+1}: {sobra['largura']:.0f} x {sobra['altura']:.0f} mm"))
                sucata_group.setLayout(sucata_layout)
                details_layout.addWidget(sucata_group)
        # --- FIM: NOVOS LABELS DE INFORMAÇÃO ---

        layout.addWidget(self.details_container)
        # --- FIM: CONTAINER PARA DETALHES ---

        cutting_widget = CuttingPlanWidget(chapa_largura, chapa_altura, self.plano, color_map)
        layout.addWidget(cutting_widget)

        buttons_layout = QHBoxLayout()
        btn_export_pdf = QPushButton("Exportar para PDF")
        btn_export_pdf.clicked.connect(self.export_to_pdf)

        # --- INÍCIO: BOTÃO PARA OCULTAR/MOSTRAR DETALHES ---
        self.toggle_details_btn = QPushButton("Ocultar Detalhes")
        self.toggle_details_btn.clicked.connect(self.toggle_details_visibility)
        # --- FIM: BOTÃO PARA OCULTAR/MOSTRAR DETALHES ---

        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(self.accept)

        buttons_layout.addWidget(self.toggle_details_btn)
        buttons_layout.addWidget(btn_export_pdf)
        buttons_layout.addStretch()
        buttons_layout.addWidget(btn_close)
        layout.addLayout(buttons_layout)

        # --- MELHORIA: Oculta os detalhes por padrão ---
        self.details_container.setVisible(False)
        self.toggle_details_btn.setText("Mostrar Detalhes")

    def toggle_details_visibility(self):
        is_visible = self.details_container.isVisible()
        self.details_container.setVisible(not is_visible)
        self.toggle_details_btn.setText("Mostrar Detalhes" if is_visible else "Ocultar Detalhes")

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
        self.resize(800, 700) # Define um tamanho inicial maior

        # Layout principal
        self.main_layout = QVBoxLayout(self)

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

        # --- Grupo de Inputs do Usuário ---
        # ... (código existente)
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
        self.status_label = QLabel("Clique em 'Calcular' para iniciar.")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-style: italic; color: #888888;")
        self.main_layout.addWidget(self.status_label)

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

        # Filtra apenas peças retangulares e agrupa por espessura
        # --- MUDANÇA: Inclui círculos no filtro ---
        valid_shapes_df = self.df[self.df['forma'].isin(['rectangle', 'circle', 'right_triangle', 'trapezoid'])].copy()
        valid_shapes_df['espessura'] = valid_shapes_df['espessura'].astype(float)
        grouped = valid_shapes_df.groupby('espessura')

        if len(grouped) == 0:
            QMessageBox.information(self, "Nenhuma Peça", "Nenhuma peça retangular, circular, triangular ou trapezoidal encontrada para o cálculo.")
            return

        # --- INÍCIO: LÓGICA DA THREAD ---
        # 1. Prepara a UI para o cálculo
        self.prepare_for_calculation()

        # 2. Cria e inicia a thread
        self.thread = CalculationThread(chapa_largura, chapa_altura, offset, grouped)
        self.thread.result_ready.connect(self.on_result_ready)
        self.thread.finished.connect(self.on_calculation_finished)
        self.thread.error.connect(self.on_calculation_error)
        self.thread.status_update.connect(self.on_status_update) # Conecta o novo sinal
        self.thread.start()
        # --- FIM: LÓGICA DA THREAD ---

    def prepare_for_calculation(self):
        """Limpa a UI e a prepara para receber novos resultados."""
        # Limpa resultados anteriores
        for i in reversed(range(self.results_scroll_layout.count())): 
            widget = self.results_scroll_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        
        self.calculation_results = {}
        self.export_report_btn.setEnabled(False)
        self.calculate_btn.setEnabled(False)
        self.calculate_btn.setText("Calculando...")
        self.status_label.setText("Calculando, por favor aguarde...")
        self.status_label.setStyleSheet("font-style: normal; color: #FFFFFF;")

    def on_result_ready(self, espessura, resultado):
        """Slot para receber o resultado de uma espessura e atualizar a UI."""
        chapa_largura = float(self.chapa_largura_input.text())
        chapa_altura = float(self.chapa_altura_input.text())
        self.calculation_results[espessura] = resultado
        self.display_results_for_thickness(espessura, resultado, chapa_largura, chapa_altura)

    def on_calculation_error(self, title, message):
        """Slot para exibir uma mensagem de erro vinda da thread."""
        QMessageBox.critical(self, title, message)

    def on_status_update(self, message):
        """Slot para receber atualizações de status da thread de cálculo."""
        self.status_label.setText(message)

    def on_calculation_finished(self):
        """Slot chamado quando todos os cálculos terminam."""
        self.calculate_btn.setEnabled(True)
        self.calculate_btn.setText("Calcular")
        self.status_label.setText("Cálculo concluído.")
        self.status_label.setStyleSheet("font-style: italic; color: #4CAF50;") # Verde
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

        # --- INÍCIO: CORREÇÃO DO ERRO "INDEX OUT OF RANGE" ---
        # Verifica se existem planos de corte antes de tentar acessá-los.
        # Isso acontece se nenhuma peça couber na chapa.
        if resultado.get('planos_unicos'):
            # --- INÍCIO: GERAÇÃO DO MAPA DE CORES ---
            # Gera o mapa de cores aqui, pois é específico para cada grupo de espessura
            offset = float(self.offset_input.text())
            tipos_de_peca_unicos = [p['tipo'] for p in resultado['planos_unicos'][0]['resumo_pecas']]
            cores = generate_distinct_colors(len(tipos_de_peca_unicos))
            self.color_map = {tipo: cor for tipo, cor in zip(tipos_de_peca_unicos, cores)}
            # --- CORREÇÃO: Armazena o mapa de cores DENTRO do resultado da espessura específica ---
            resultado['color_map'] = self.color_map
            # --- FIM: GERAÇÃO DO MAPA DE CORES ---

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
        else:
            # Caso não haja planos, exibe uma mensagem informativa.
            no_fit_label = QLabel("Nenhuma peça coube na chapa com as dimensões fornecidas.")
            no_fit_label.setStyleSheet("color: #FDBA74; font-style: italic;") # Laranja claro
            group_layout.addWidget(no_fit_label)
        # --- FIM: CORREÇÃO ---

        group_box.setLayout(group_layout)
        self.results_scroll_layout.addWidget(group_box)

    def show_plan_visualization(self, plano_info, chapa_w, chapa_h, color_map):
        offset = float(self.offset_input.text())
        dialog = PlanVisualizationDialog(chapa_w, chapa_h, plano_info, offset, color_map, self)
        dialog.exec_()