# pdf_generator.py

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# =============================================================================
# CONSTANTES GLOBAIS DE LAYOUT DA PÁGINA
# =============================================================================
PAGE_WIDTH, PAGE_HEIGHT = A4
MARGEM_GERAL = 15 * mm
HEADER_AREA_ALTURA = 25 * mm
FOOTER_AREA_ALTURA = 25 * mm 

# FATOR DE ESCALA PARA FONTES (para facilitar futuras manutenções)
# Aumenta o tamanho de todas as fontes nos desenhos técnicos em 20%.
FONT_SCALE_FACTOR = 1.2

# =============================================================================
# FUNÇÕES UTILITÁRIAS E DE DESENHO DE COMPONENTES
# Estas devem ser definidas ANTES das funções que as usam.
# =============================================================================

def formatar_numero(valor):
    """Formata um número para string, removendo o decimal se for zero."""
    if valor is None:
        return "0"
    if valor == int(valor):
        return str(int(valor))
    return str(valor).replace('.', ',')

def desenhar_cabecalho(c, nome_arquivo):
    c.setFont("Helvetica-Bold", 14 * FONT_SCALE_FACTOR)
    y_pos_texto = PAGE_HEIGHT - HEADER_AREA_ALTURA + (10 * mm)
    c.drawCentredString(PAGE_WIDTH / 2, y_pos_texto, f"Desenho da Peça: {nome_arquivo}")
    y_pos_linha = PAGE_HEIGHT - HEADER_AREA_ALTURA
    c.line(MARGEM_GERAL, y_pos_linha, PAGE_WIDTH - MARGEM_GERAL, y_pos_linha)

def desenhar_rodape_aprimorado(c, row):
    largura_total = PAGE_WIDTH - 2 * MARGEM_GERAL
    altura_bloco = 12 * mm
    y_rodape = FOOTER_AREA_ALTURA - altura_bloco - (5 * mm)
    c.setStrokeColorRGB(0, 0, 0)
    c.rect(MARGEM_GERAL, y_rodape, largura_total, altura_bloco)
    
    coluna1_x, coluna2_x, coluna3_x = MARGEM_GERAL, MARGEM_GERAL + largura_total * 0.60, MARGEM_GERAL + largura_total * 0.80
    c.line(coluna2_x, y_rodape, coluna2_x, y_rodape + altura_bloco)
    c.line(coluna3_x, y_rodape, coluna3_x, y_rodape + altura_bloco)
    
    y_titulo, y_valor = y_rodape + altura_bloco - 4.5*mm, y_rodape + 3.5*mm
    
    c.setFont("Helvetica", 7 * FONT_SCALE_FACTOR)
    c.drawCentredString(coluna1_x + (coluna2_x - coluna1_x)/2, y_titulo, "NOME DA PEÇA / IDENTIFICADOR")
    c.setFont("Helvetica-Bold", 10 * FONT_SCALE_FACTOR)
    c.drawCentredString(coluna1_x + (coluna2_x - coluna1_x)/2, y_valor, str(row.get('nome_arquivo', 'N/A')))
    
    c.setFont("Helvetica", 7 * FONT_SCALE_FACTOR)
    c.drawCentredString(coluna2_x + (coluna3_x - coluna2_x)/2, y_titulo, "ESPESSURA")
    c.setFont("Helvetica-Bold", 10 * FONT_SCALE_FACTOR)
    c.drawCentredString(coluna2_x + (coluna3_x - coluna2_x)/2, y_valor, f"{formatar_numero(row.get('espessura', 0))} mm")
    
    c.setFont("Helvetica", 7 * FONT_SCALE_FACTOR)
    c.drawCentredString(coluna3_x + (PAGE_WIDTH - MARGEM_GERAL - coluna3_x)/2, y_titulo, "QUANTIDADE")
    c.setFont("Helvetica-Bold", 10 * FONT_SCALE_FACTOR)
    c.drawCentredString(coluna3_x + (PAGE_WIDTH - MARGEM_GERAL - coluna3_x)/2, y_valor, formatar_numero(row.get('qtd', 0)))

def desenhar_erro_dados(c, forma):
    c.setFont("Helvetica-Bold", 14 * FONT_SCALE_FACTOR)
    c.drawCentredString(A4[0]/2, A4[1]/2, f"Dados inválidos para a forma: '{forma}'")

def desenhar_cota_horizontal(c, x1, x2, y, texto):
    FONT_NAME, FONT_SIZE = "Helvetica", 10 * FONT_SCALE_FACTOR
    c.setFont(FONT_NAME, FONT_SIZE)
    c.line(x1, y, x2, y)
    tick_len = 2 * mm
    c.line(x1, y - tick_len/2, x1 + tick_len/2, y + tick_len/2)
    c.line(x2, y - tick_len/2, x2 - tick_len/2, y + tick_len/2)
    largura_texto = c.stringWidth(texto, FONT_NAME, FONT_SIZE)
    centro_x, y_texto = (x1 + x2) / 2, y + 1*mm
    padding = 1.5 * mm
    gap_inicio, gap_fim = centro_x - (largura_texto / 2) - padding, centro_x + (largura_texto / 2) + padding
    c.line(x1, y, gap_inicio, y)
    c.line(gap_fim, y, x2, y)
    c.drawCentredString(centro_x, y_texto, texto)

def desenhar_cota_vertical(c, y1, y2, x, texto):
    FONT_NAME, FONT_SIZE = "Helvetica", 10 * FONT_SCALE_FACTOR
    c.setFont(FONT_NAME, FONT_SIZE)
    c.line(x, y1, x, y2)
    tick_len = 2 * mm
    c.line(x - tick_len/2, y1, x + tick_len/2, y1 + tick_len/2)
    c.line(x - tick_len/2, y2, x + tick_len/2, y2 - tick_len/2)
    largura_texto, altura_texto = c.stringWidth(texto, FONT_NAME, FONT_SIZE), FONT_SIZE
    centro_y = (y1 + y2) / 2
    c.saveState()
    c.translate(x, centro_y)
    c.rotate(90)
    c.setFillColorRGB(1, 1, 1) # Cor de fundo do "buraco" no meio da linha de cota
    c.rect(-largura_texto/2 - 1*mm, -altura_texto/2, largura_texto + 2*mm, altura_texto, stroke=0, fill=1)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(0, -1.5*mm, texto)
    c.restoreState()

def desenhar_cota_diametro_furo(c, x, y, raio, diametro_real):
    c.saveState()
    c.setFont("Helvetica", 10 * FONT_SCALE_FACTOR)
    texto = f"Ø {formatar_numero(diametro_real)}"
    largura_texto = c.stringWidth(texto, "Helvetica", 10 * FONT_SCALE_FACTOR)
    p_borda_x, p_borda_y = x + raio * 0.7071, y + raio * 0.7071 # Ponto a 45 graus na borda do círculo
    p_meio_x, p_meio_y = p_borda_x + 4 * mm, p_borda_y + 4 * mm
    p_final_x = p_meio_x + largura_texto + 2*mm
    path = c.beginPath()
    path.moveTo(p_borda_x, p_borda_y)
    path.lineTo(p_meio_x, p_meio_y)
    path.lineTo(p_final_x, p_meio_y)
    c.drawPath(path)
    c.drawString(p_meio_x + 1*mm, p_meio_y + 1*mm, texto)
    c.restoreState()

# =============================================================================
# FUNÇÕES DE DESENHO PARA CADA FORMA GEOMÉTRICA
# =============================================================================

def desenhar_retangulo(c, row):
    largura, altura = row.get('largura', 0), row.get('altura', 0)
    if largura <= 0 or altura <= 0: desenhar_erro_dados(c, "Retângulo"); return
    
    max_w, max_h = PAGE_WIDTH - 2*MARGEM_GERAL, PAGE_HEIGHT - HEADER_AREA_ALTURA - FOOTER_AREA_ALTURA
    dist_cota_furo, dist_cota_total, overshoot = 8*mm, 16*mm, 2*mm
    espaco_cota_x, espaco_cota_y = dist_cota_total, dist_cota_total
    
    escala = min((max_w - espaco_cota_x) / largura, (max_h - espaco_cota_y) / altura) * 0.95
    dw, dh = largura*escala, altura*escala
    bloco_visual_width, bloco_visual_height = dw + espaco_cota_x, dh + espaco_cota_y
    
    base_desenho_y = FOOTER_AREA_ALTURA
    inicio_bloco_x, inicio_bloco_y = MARGEM_GERAL + (max_w - bloco_visual_width)/2, base_desenho_y + (max_h - bloco_visual_height)/2
    x0, y0 = inicio_bloco_x + espaco_cota_x, inicio_bloco_y + espaco_cota_y
    c.rect(x0, y0, dw, dh)
    
    furos = row.get('furos')
    if isinstance(furos, list) and furos:
        for furo in furos: c.circle(x0 + (furo['x']*escala), y0 + (furo['y']*escala), (furo['diam']/2)*escala, stroke=1, fill=0)
        y_pos_cota = y0 - dist_cota_furo
        unique_x = sorted(list(set(f['x'] for f in furos)))
        dim_points_x = [0] + unique_x + [largura]
        for x_real in dim_points_x: c.line(x0 + x_real*escala, y0, x0 + x_real*escala, y_pos_cota - overshoot)
        for i in range(1, len(dim_points_x)):
            start_x, end_x = dim_points_x[i-1], dim_points_x[i]
            if (end_x - start_x)>0.01: desenhar_cota_horizontal(c, x0+start_x*escala, x0+end_x*escala, y_pos_cota, formatar_numero(end_x-start_x))
        x_pos_cota = x0 - dist_cota_furo
        unique_y = sorted(list(set(f['y'] for f in furos)))
        dim_points_y = [0] + unique_y + [altura]
        for y_real in dim_points_y: c.line(x0, y0 + y_real*escala, x_pos_cota - overshoot, y0 + y_real*escala)
        for i in range(1, len(dim_points_y)):
            start_y, end_y = dim_points_y[i-1], dim_points_y[i]
            if (end_y - start_y)>0.01: desenhar_cota_vertical(c, y0+start_y*escala, y0+end_y*escala, x_pos_cota, formatar_numero(end_y-start_y))
        # --- CORREÇÃO: Só desenha a cota do furo se a lista de furos não estiver vazia ---
        if furos:
            desenhar_cota_diametro_furo(c, x0+furos[0]['x']*escala, y0+furos[0]['y']*escala, (furos[0]['diam']/2)*escala, furos[0]['diam'])
        
    y_cota_total = y0 - dist_cota_total
    c.line(x0, y0, x0, y_cota_total-overshoot)
    c.line(x0+dw, y0, x0+dw, y_cota_total-overshoot)
    desenhar_cota_horizontal(c, x0, x0+dw, y_cota_total, formatar_numero(largura))
    
    x_cota_total = x0 - dist_cota_total
    c.line(x0, y0, x_cota_total-overshoot, y0)
    c.line(x0, y0+dh, x_cota_total-overshoot, y0+dh)
    desenhar_cota_vertical(c, y0, y0+dh, x_cota_total, formatar_numero(altura))

def desenhar_circulo(c, row):
    diametro = row.get('diametro', 0)
    if diametro <= 0: desenhar_erro_dados(c, "Círculo"); return
    
    max_w, max_h = PAGE_WIDTH-2*MARGEM_GERAL, PAGE_HEIGHT-HEADER_AREA_ALTURA-FOOTER_AREA_ALTURA
    dist_cota, overshoot = 8*mm, 2*mm
    escala = min((max_w-dist_cota*2)/diametro, (max_h-dist_cota*2)/diametro)*0.95
    raio_desenhado = (diametro*escala)/2
    bloco_visual_size = (raio_desenhado*2)+dist_cota*2
    base_desenho_y = FOOTER_AREA_ALTURA
    inicio_bloco_x, inicio_bloco_y = MARGEM_GERAL + (max_w-bloco_visual_size)/2, base_desenho_y + (max_h-bloco_visual_size)/2
    cx, cy = inicio_bloco_x+dist_cota+raio_desenhado, inicio_bloco_y+dist_cota+raio_desenhado
    c.circle(cx, cy, raio_desenhado, stroke=1, fill=0)
    
    furos = row.get('furos')
    if isinstance(furos, list) and furos:
        x0_peca, y0_peca = cx-raio_desenhado, cy-raio_desenhado
        for furo in furos: c.circle(x0_peca+(furo['x']*escala), y0_peca+(furo['y']*escala), (furo['diam']/2)*escala, stroke=1, fill=0)
        
    y_cota_h = cy-raio_desenhado-dist_cota
    c.line(cx-raio_desenhado, cy-raio_desenhado, cx-raio_desenhado, y_cota_h-overshoot)
    c.line(cx+raio_desenhado, cy-raio_desenhado, cx+raio_desenhado, y_cota_h-overshoot)
    desenhar_cota_horizontal(c, cx-raio_desenhado, cx+raio_desenhado, y_cota_h, f"Ø {formatar_numero(diametro)}")

def desenhar_triangulo_retangulo(c, row):
    base, altura = row.get('rt_base', 0), row.get('rt_height', 0)
    if base <= 0 or altura <= 0: desenhar_erro_dados(c, "Triângulo Retângulo"); return
    
    max_w, max_h = PAGE_WIDTH-2*MARGEM_GERAL, PAGE_HEIGHT-HEADER_AREA_ALTURA-FOOTER_AREA_ALTURA
    dist_cota, overshoot = 8*mm, 2*mm
    espaco_cota_x, espaco_cota_y = dist_cota, dist_cota
    escala = min((max_w-espaco_cota_x)/base, (max_h-espaco_cota_y)/altura)*0.95
    db, dh = base*escala, altura*escala
    bloco_visual_width, bloco_visual_height = db+espaco_cota_x, dh+espaco_cota_y
    base_desenho_y = FOOTER_AREA_ALTURA
    inicio_bloco_x, inicio_bloco_y = MARGEM_GERAL+(max_w-bloco_visual_width)/2, base_desenho_y+(max_h-bloco_visual_height)/2
    x0, y0 = inicio_bloco_x+espaco_cota_x, inicio_bloco_y+espaco_cota_y
    path=c.beginPath()
    path.moveTo(x0, y0)
    path.lineTo(x0+db, y0)
    path.lineTo(x0, y0+dh)
    path.close()
    c.drawPath(path)
    
    furos = row.get('furos')
    if isinstance(furos, list) and furos:
        for furo in furos: c.circle(x0+(furo['x']*escala), y0+(furo['y']*escala), (furo['diam']/2)*escala, stroke=1, fill=0)
        
    y_cota_h = y0 - dist_cota
    c.line(x0, y0, x0, y_cota_h-overshoot)
    c.line(x0+db, y0, x0+db, y_cota_h-overshoot)
    desenhar_cota_horizontal(c, x0, x0+db, y_cota_h, formatar_numero(base))
    
    x_cota_v = x0 - dist_cota
    c.line(x0, y0, x_cota_v-overshoot, y0)
    c.line(x0, y0+dh, x_cota_v-overshoot, y0+dh)
    desenhar_cota_vertical(c, y0, y0+dh, x_cota_v, formatar_numero(altura))

def desenhar_trapezio(c, row):
    large_base, small_base, height = row.get('trapezoid_large_base', 0), row.get('trapezoid_small_base', 0), row.get('trapezoid_height', 0)
    if large_base<=0 or height<=0 or small_base<=0: desenhar_erro_dados(c, "Trapézio"); return
    
    max_w, max_h = PAGE_WIDTH-2*MARGEM_GERAL, PAGE_HEIGHT-HEADER_AREA_ALTURA-FOOTER_AREA_ALTURA
    dist_cota, overshoot = 8*mm, 2*mm
    espaco_cota_x, espaco_cota_y = dist_cota, dist_cota*2
    escala = min((max_w-espaco_cota_x)/large_base, (max_h-espaco_cota_y)/height)*0.95
    dlb, dsb, dh = large_base*escala, small_base*escala, height*escala
    bloco_visual_width, bloco_visual_height = dlb+espaco_cota_x, dh+espaco_cota_y
    base_desenho_y = FOOTER_AREA_ALTURA
    inicio_bloco_x, inicio_bloco_y = MARGEM_GERAL+(max_w-bloco_visual_width)/2, base_desenho_y+(max_h-bloco_visual_height)/2
    x0, y0 = inicio_bloco_x+espaco_cota_x/2, inicio_bloco_y+dist_cota
    x_offset = (dlb - dsb)/2
    p1, p2, p3, p4 = (x0, y0), (x0+dlb, y0), (x0+dlb-x_offset, y0+dh), (x0+x_offset, y0+dh)
    path = c.beginPath()
    path.moveTo(*p1)
    path.lineTo(*p2)
    path.lineTo(*p3)
    path.lineTo(*p4)
    path.close()
    c.drawPath(path)
    
    furos = row.get('furos')
    if isinstance(furos, list) and furos:
        for furo in furos: c.circle(x0+(furo['x']*escala), y0+(furo['y']*escala), (furo['diam']/2)*escala, stroke=1, fill=0)
        
    y_cota_h1 = y0 - dist_cota
    c.line(p1[0], p1[1], p1[0], y_cota_h1-overshoot)
    c.line(p2[0], p2[1], p2[0], y_cota_h1-overshoot)
    desenhar_cota_horizontal(c, p1[0], p2[0], y_cota_h1, formatar_numero(large_base))
    
    y_cota_h2 = p3[1]+dist_cota
    c.line(p4[0], p4[1], p4[0], y_cota_h2+overshoot)
    c.line(p3[0], p3[1], p3[0], y_cota_h2+overshoot)
    desenhar_cota_horizontal(c, p4[0], p3[0], y_cota_h2, formatar_numero(small_base))
    
    x_cota_v = min(p1[0], p4[0]) - dist_cota
    c.line(p1[0], p1[1], x_cota_v-overshoot, p1[1])
    c.line(p4[0], p4[1], x_cota_v-overshoot, p4[1])
    desenhar_cota_vertical(c, p1[1], p4[1], x_cota_v, formatar_numero(height))
color_map = {
            'tipo1': (0.2, 0.6, 0.2),  # Verde
            'tipo2': (0.2, 0.2, 0.7),  # Azul
            'tipo3': (0.7, 0.2, 0.2),  # Vermelho
            'tipo4': (0.6, 0.6, 0.2),  # Amarelo
            'tipo5': (0.6, 0.2, 0.6),  # Roxo
            'tipo6': (0.2, 0.6, 0.6),  # Ciano

            # Adicione mais tipos e cores conforme necessário
        }
def gerar_pdf_plano_de_corte(c, chapa_largura, chapa_altura, plano, color_map_qt=None):
    """
    Desenha um plano de corte (nesting) em uma página de PDF.
    
    :param c: O canvas do reportlab.
    :param chapa_largura: Largura real da chapa.
    :param chapa_altura: Altura real da chapa.
    :param plano: Lista de dicionários de peças, cada um com 'x', 'y', 'largura', 'altura'.
    :param color_map_qt: Opcional. Dicionário de QColor para mapear cores da UI.
    """
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 20*mm, "Plano de Corte da Chapa")

    max_w, max_h = PAGE_WIDTH - 2*MARGEM_GERAL, PAGE_HEIGHT - 40*mm # Mais espaço para título
    dist_cota, overshoot = 8*mm, 2*mm
    espaco_cota_x, espaco_cota_y = dist_cota, dist_cota

    # Calcula a escala para caber na página
    escala = min((max_w - espaco_cota_x) / chapa_largura, (max_h - espaco_cota_y) / chapa_altura) * 0.95
    dw, dh = chapa_largura * escala, chapa_altura * escala
    
    # Centraliza o desenho
    bloco_visual_width, bloco_visual_height = dw + espaco_cota_x, dh + espaco_cota_y
    inicio_bloco_x = MARGEM_GERAL + (max_w - bloco_visual_width) / 2
    inicio_bloco_y = MARGEM_GERAL + (max_h - bloco_visual_height) / 2
    x0, y0 = inicio_bloco_x + espaco_cota_x, inicio_bloco_y + espaco_cota_y

    # 1. Desenha a chapa
    c.setStrokeColorRGB(0.8, 0.8, 0.8) # Cinza claro para a chapa
    c.rect(x0, y0, dw, dh, stroke=1, fill=0)
    c.setStrokeColorRGB(0, 0, 0) # Volta para preto

    # 2. Desenha as peças
    default_color = (0.66, 0.26, 0.26) # Vermelho escuro padrão

    for peca in plano:
        w, h, x, y, tipo_key = peca['largura'], peca['altura'], peca['x'], peca['y'], peca['tipo_key']
        rect_w, rect_h = w * escala, h * escala
        rect_x = x0 + x * escala
        # A coordenada Y da peça já está invertida (origem no topo).
        # Para desenhar no PDF (origem embaixo), calculamos o canto inferior esquerdo.
        rect_y = (y0 + dh) - (y * escala) - rect_h

        # Define a cor da peça

        if color_map_qt and tipo_key in color_map_qt:
            q_color = color_map_qt[tipo_key]
            c.setFillColorRGB(q_color.redF(), q_color.greenF(), q_color.blueF()) # Converte de QColor para reportlab
        else:
            # Fallback para o mapa de cores interno se o mapa da UI não for fornecido ou não tiver a chave
            rgb_color = color_map.get(tipo_key, default_color)
            c.setFillColorRGB(*default_color)

        c.rect(rect_x, rect_y, rect_w, rect_h, stroke=1, fill=1)

        # 3. Desenha os furos
        furos = peca.get('furos', [])
        if furos:
            c.setFillColorRGB(1, 1, 1) # Furos brancos
            for furo in furos:
                furo_x_centro = rect_x + furo['x'] * escala
                furo_y_centro = rect_y + rect_h - (furo['y'] * escala) # Y do furo é relativo ao topo da peça
                c.circle(furo_x_centro, furo_y_centro, (furo['diam'] / 2) * escala, stroke=1, fill=1)

    # 3. Desenha as cotas da chapa
    y_cota_total = y0 - dist_cota
    desenhar_cota_horizontal(c, x0, x0 + dw, y_cota_total, formatar_numero(chapa_largura))
    x_cota_total = x0 - dist_cota
    desenhar_cota_vertical(c, y0, y0 + dh, x_cota_total, formatar_numero(chapa_altura))

def _desenhar_plano_unico_com_detalhes(c, y_start, plano_info, chapa_largura, chapa_altura, plano_idx, color_map):
    """
    Função auxiliar para desenhar um único plano de corte com seus detalhes em uma área específica da página.
    Retorna a posição Y final após o desenho.
    """
    c.setFont("Helvetica-Bold", 11)
    titulo = f"Plano de Corte {plano_idx + 1} (Repetir {plano_info['repeticoes']}x)"
    c.drawString(MARGEM_GERAL, y_start, titulo)
    y_cursor = y_start - 5*mm

    # Área de desenho do plano (à esquerda)
    area_desenho_w = (PAGE_WIDTH / 2) - (1.5 * MARGEM_GERAL)
    area_desenho_h = 100 * mm # Altura fixa para a visualização

    # Escala e centralização
    escala = min(area_desenho_w / chapa_largura, area_desenho_h / chapa_altura) * 0.95
    dw, dh = chapa_largura * escala, chapa_altura * escala
    
    # A coordenada Y do desenho começa no topo da área de desenho
    x_origem_desenho = MARGEM_GERAL + (area_desenho_w - dw) / 2
    y_origem_desenho = y_cursor - (area_desenho_h - dh) / 2

    # Desenha chapa e peças
    c.setStrokeColorRGB(0.8, 0.8, 0.8)
    # O rect é desenhado de (x, y) com largura e altura, então o y precisa ser o canto inferior esquerdo.
    c.rect(x_origem_desenho, y_origem_desenho - dh, dw, dh, stroke=1, fill=0)
    c.setStrokeColorRGB(0, 0, 0)
    
    # --- INÍCIO: CORREÇÃO DE CORES DAS PEÇAS ---
    for peca in plano_info['plano']:
        # Define a cor da peça usando o color_map
        q_color = color_map.get(peca['tipo_key'])
        if q_color:
            c.setFillColorRGB(q_color.redF(), q_color.greenF(), q_color.blueF())
        else:
            c.setFillColorRGB(0.66, 0.26, 0.26) # Cor padrão de fallback
        
        # --- INÍCIO: DESENHO CONDICIONAL (RETÂNGULO OU CÍRCULO) ---
        rect_y_inferior = y_origem_desenho - (peca['y'] * escala) - (peca['altura'] * escala)
        if peca.get('forma') == 'circle':
            diametro_original = peca.get('diametro', 0)
            raio_desenhado = (diametro_original * escala) / 2
            centro_x = (x_origem_desenho + peca['x'] * escala) + (peca['largura'] * escala) / 2
            centro_y = rect_y_inferior + (peca['altura'] * escala) / 2
            c.circle(centro_x, centro_y, raio_desenhado, stroke=1, fill=1)
        elif peca.get('forma') == 'paired_triangle':
            # Desenha dois triângulos opostos
            x, y, w, h = x_origem_desenho + peca['x'] * escala, rect_y_inferior, peca['largura'] * escala, peca['altura'] * escala
            path = c.beginPath()
            path.moveTo(x, y)
            path.lineTo(x + w, y)
            path.lineTo(x, y + h)
            path.close()
            c.drawPath(path, stroke=1, fill=1)
            path = c.beginPath()
            path.moveTo(x + w, y + h)
            path.lineTo(x, y + h)
            path.lineTo(x + w, y)
            path.close()
            c.drawPath(path, stroke=1, fill=1)
        else: # 'rectangle'
            c.rect(x_origem_desenho + peca['x'] * escala, rect_y_inferior, peca['largura'] * escala, peca['altura'] * escala, stroke=1, fill=1)
        # --- FIM: DESENHO CONDICIONAL ---

    # Desenha as sobras
    sobras = plano_info.get('sobras', [])
    if sobras:
        for sobra in sobras:
            if sobra.get('tipo_sobra') == 'aproveitavel':
                c.setFillColorRGB(0.39, 0.39, 0.39, 0.6) # Cinza escuro semi-transparente
                c.setStrokeColorRGB(0.2, 0.2, 0.2)
            else:
                c.setFillColorRGB(0.9, 0.9, 0.9, 0.5) # Cinza claro semi-transparente
                c.setStrokeColorRGB(0.5, 0.5, 0.5)
            sobra_y_inferior = y_origem_desenho - (sobra['y'] * escala) - (sobra['altura'] * escala)
            c.rect(x_origem_desenho + sobra['x'] * escala, sobra_y_inferior, sobra['largura'] * escala, sobra['altura'] * escala, stroke=1, fill=1)


    # Área da lista de peças (à direita)
    x_lista = (PAGE_WIDTH / 2) + (MARGEM_GERAL / 2)
    y_lista = y_cursor
    # --- CORREÇÃO: Garante que o texto seja preto ---
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_lista, y_lista, "Peças neste plano:")
    y_lista -= 4*mm
    c.setFont("Helvetica", 8)
    for item in plano_info['resumo_pecas']:
        # Verifica se precisa pular para a próxima linha na lista
        if y_lista < (y_cursor - area_desenho_h + 5*mm):
            c.drawString(x_lista, y_lista, "...")
            break
        texto_peca = f"- {item['qtd']}x de {item['tipo']} mm"
        c.drawString(x_lista, y_lista, texto_peca)
        y_lista -= 3.5*mm

    # --- INÍCIO: ADIÇÃO DO RESUMO DE SOBRAS NO PDF ---
    y_lista -= 5*mm # Espaço antes das sobras
    sobras_aproveitaveis = [s for s in sobras if s.get('tipo_sobra') == 'aproveitavel']
    sobras_sucata = [s for s in sobras if s.get('tipo_sobra') != 'aproveitavel']

    if sobras_aproveitaveis:
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x_lista, y_lista, "Sobras Aproveitáveis:")
        y_lista -= 4*mm
        c.setFont("Helvetica", 8)
        for sobra in sobras_aproveitaveis:
            c.drawString(x_lista, y_lista, f"- {sobra['largura']:.0f} x {sobra['altura']:.0f} mm")
            y_lista -= 3.5*mm
        y_lista -= 2*mm # Espaço extra

    if sobras_sucata:
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x_lista, y_lista, "Sobras (Sucata):")
        y_lista -= 4*mm
        c.setFont("Helvetica", 8)
        for sobra in sobras_sucata:
            c.drawString(x_lista, y_lista, f"- {sobra['largura']:.0f} x {sobra['altura']:.0f} mm")
            y_lista -= 3.5*mm
    # --- FIM: ADIÇÃO DO RESUMO DE SOBRAS NO PDF ---

    return y_cursor - area_desenho_h - 5*mm # Retorna a nova posição Y

def gerar_pdf_aproveitamento_completo(c, resultados_completos, chapa_largura, chapa_altura):
    """
    Gera um relatório PDF completo com todos os planos de corte para todas as espessuras.
    """
    # --- INÍCIO: MELHORIA NO DESIGN DO CABEÇALHO ---
    c.saveState()
    c.setFillColorRGB(0.13, 0.13, 0.13) # Fundo cinza escuro
    c.rect(0, PAGE_HEIGHT - 25*mm, PAGE_WIDTH, 25*mm, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1) # Texto branco
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 12*mm, "Relatório de Aproveitamento de Chapa")
    c.setFont("Helvetica", 11)
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 19*mm, f"Dimensões da Chapa: {formatar_numero(chapa_largura)} x {formatar_numero(chapa_altura)} mm")
    c.restoreState()
    # --- FIM: MELHORIA NO DESIGN DO CABEÇALHO ---

    y_cursor = PAGE_HEIGHT - MARGEM_GERAL - 20*mm

    for espessura, resultado in resultados_completos.items():
        # Verifica se há espaço para o cabeçalho da espessura
        if y_cursor < MARGEM_GERAL + 40*mm:
            c.showPage()
            y_cursor = PAGE_HEIGHT - MARGEM_GERAL

        # Cabeçalho da Espessura
        c.saveState()
        c.setFillColorRGB(0.9, 0.9, 0.9) # Fundo cinza claro
        c.setStrokeColorRGB(0,0,0)
        c.rect(MARGEM_GERAL, y_cursor - 5*mm, PAGE_WIDTH - 2*MARGEM_GERAL, 7*mm, stroke=1, fill=1)
        c.setFillColorRGB(0,0,0) # Texto preto
        c.setFont("Helvetica-Bold", 12)
        c.drawString(MARGEM_GERAL + 3*mm, y_cursor - 3.5*mm, f"Espessura: {espessura} mm")
        c.restoreState()
        y_cursor -= 8*mm

        # --- CORREÇÃO: Garante que o texto seja preto ---
        c.setFillColorRGB(0, 0, 0)
        # Sumário da Espessura
        c.setFont("Helvetica", 10)
        sumario = f"Total de Chapas: {resultado['total_chapas']}   |   Aproveitamento Geral: {resultado['aproveitamento_geral']}"
        c.drawString(MARGEM_GERAL, y_cursor, sumario)
        
        # --- INÍCIO: CÁLCULO E EXIBIÇÃO DO PESO TOTAL DAS CHAPAS ---
        peso_total_chapas_kg = (chapa_largura / 1000) * (chapa_altura / 1000) * espessura * 7.85 * resultado['total_chapas']
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(PAGE_WIDTH - MARGEM_GERAL, y_cursor, f"Peso Total das Chapas: {peso_total_chapas_kg:.2f} kg")
        # --- FIM: CÁLCULO E EXIBIÇÃO DO PESO TOTAL DAS CHAPAS ---
        y_cursor -= 8*mm

        # --- INÍCIO: RESUMO DE PESO DAS SOBRAS ---
        from collections import Counter
        todas_as_sobras = []
        for plano in resultado['planos_unicos']:
            todas_as_sobras.extend(plano.get('sobras', []) * plano['repeticoes'])
        
        if todas_as_sobras:
            sobras_aproveitaveis = [s for s in todas_as_sobras if s.get('tipo_sobra') == 'aproveitavel']
            sobras_sucata = [s for s in todas_as_sobras if s.get('tipo_sobra') != 'aproveitavel']
            area_aproveitavel_mm2 = sum(s['largura'] * s['altura'] for s in sobras_aproveitaveis)
            area_sucata_mm2 = sum(s['largura'] * s['altura'] for s in sobras_sucata)
            peso_aproveitavel_kg = (area_aproveitavel_mm2 / 1_000_000) * espessura * 7.85
            peso_sucata_kg = (area_sucata_mm2 / 1_000_000) * espessura * 7.85
            
            c.setFont("Helvetica-Bold", 9)
            c.drawString(MARGEM_GERAL, y_cursor, f"Peso Sobras Aproveitáveis: {peso_aproveitavel_kg:.2f} kg")
            c.drawString(MARGEM_GERAL + 70*mm, y_cursor, f"Peso Sobras (Sucata): {peso_sucata_kg:.2f} kg")
            y_cursor -= 5*mm
        # --- FIM: RESUMO DE PESO DAS SOBRAS ---

        # Desenha cada plano de corte único
        for i, plano_info in enumerate(resultado['planos_unicos']):
            # Verifica se o plano cabe na página atual, se não, cria uma nova
            altura_necessaria_plano = 115 * mm # Altura estimada para cada bloco de plano
            if y_cursor < altura_necessaria_plano:
                c.showPage()
                y_cursor = PAGE_HEIGHT - MARGEM_GERAL
                # Adiciona um cabeçalho de continuação
                c.setFont("Helvetica-Bold", 12)
                c.drawString(MARGEM_GERAL, y_cursor, f"Continuação - Espessura: {espessura} mm")
                y_cursor -= 10*mm

            y_cursor = _desenhar_plano_unico_com_detalhes(c, y_cursor, plano_info, chapa_largura, chapa_altura, i, resultado['color_map'])
            
            # Linha separadora
            if i < len(resultado['planos_unicos']) - 1:
                c.setStrokeColorRGB(0.8, 0.8, 0.8)
                c.line(MARGEM_GERAL, y_cursor, PAGE_WIDTH - MARGEM_GERAL, y_cursor)
                y_cursor -= 5*mm

        # Espaço antes da próxima espessura
        y_cursor -= 10*mm

# =============================================================================
# FUNÇÃO PRINCIPAL DO MÓDULO (INTERFACE PÚBLICA)
# Esta função chama as outras, por isso deve vir por último.
# =============================================================================

def desenhar_forma(c, row):
    """
    Função principal que desenha o cabeçalho, rodapé e a forma geométrica correta.
    """
    desenhar_cabecalho(c, row.get('nome_arquivo', 'SEM NOME'))
    desenhar_rodape_aprimorado(c, row)
    
    forma = str(row.get('forma', '')).strip().lower()
    
    if forma == 'rectangle':
        desenhar_retangulo(c, row)
    elif forma == 'circle':
        desenhar_circulo(c, row)
    elif forma == 'right_triangle':
        desenhar_triangulo_retangulo(c, row)
    elif forma == 'trapezoid':
        desenhar_trapezio(c, row)
    else:
        c.setFont("Helvetica", 12)
        c.drawCentredString(A4[0]/2, A4[1]/2, f"Forma '{forma}' desconhecida ou não implementada.")
