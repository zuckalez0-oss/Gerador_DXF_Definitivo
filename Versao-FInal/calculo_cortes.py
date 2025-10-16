import rectpack
import logging
from PyQt5.QtCore import QObject, pyqtSignal
from rectpack.maxrects import MaxRectsBssf, MaxRectsBaf, MaxRectsBlsf, MaxRectsBl
from rectpack.skyline import SkylineBl, SkylineBlWm, SkylineMwf, SkylineMwfl
import math
import os

# --- INÍCIO: CONFIGURAÇÃO DE LOGGING PARA DEBUG ---
logging.basicConfig(filename='debug_nesting.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s', filemode='w')
# --- FIM: CONFIGURAÇÃO DE LOGGING ---

# --- INÍCIO: CLASSE PARA EMITIR SINAIS DE STATUS ---
class StatusSignaler(QObject):
    status_update = pyqtSignal(str)

status_signaler = StatusSignaler()
# --- FIM: CLASSE PARA EMITIR SINAIS DE STATUS ---

def _merge_scraps(scraps):
    """
    Função auxiliar para fundir retângulos de sobras adjacentes.
    Esta versão é mais robusta e corrige falhas na lógica de fusão horizontal.
    """
    if not scraps:
        return []

    TOLERANCE = 1e-5 

    while True:
        merged_in_pass = False
        i = 0
        while i < len(scraps):
            j = i + 1
            while j < len(scraps):
                r1 = scraps[i]
                r2 = scraps[j]
                
                merged = False

                # Tenta fusão vertical
                if abs(r1['x'] - r2['x']) < TOLERANCE and abs(r1['largura'] - r2['largura']) < TOLERANCE:
                    if abs((r1['y'] + r1['altura']) - r2['y']) < TOLERANCE:
                        r1['altura'] += r2['altura']
                        merged = True
                    elif abs((r2['y'] + r2['altura']) - r1['y']) < TOLERANCE:
                        r1['y'] = r2['y']
                        r1['altura'] += r2['altura']
                        merged = True

                # Tenta fusão horizontal
                elif abs(r1['y'] - r2['y']) < TOLERANCE and abs(r1['altura'] - r2['altura']) < TOLERANCE:
                    if abs((r1['x'] + r1['largura']) - r2['x']) < TOLERANCE:
                        r1['largura'] += r2['largura']
                        merged = True
                    elif abs((r2['x'] + r2['largura']) - r1['x']) < TOLERANCE:
                        # r1 está à direita de r2, expande r1 para a esquerda
                        r1['x'] = r2['x']
                        r1['largura'] += r2['largura']
                        merged = True

                if merged:
                    scraps.pop(j)
                    merged_in_pass = True
                    j = i + 1 
                else:
                    j += 1
            i += 1
        
        if not merged_in_pass:
            break
            
    return scraps

def encontrar_sobras(chapa_largura, chapa_altura, pecas_alocadas, min_dim=50):
    """
    Encontra os maiores retângulos de sobra em uma chapa.
    Este método usa um algoritmo de varredura (Sweep-line) para robustez,
    seguido por uma etapa de fusão para consolidar os resultados.
    """
    logging.debug("Iniciando 'encontrar_sobras' com algoritmo de varredura (Scanline).")

    # 1. Coletar todas as coordenadas Y únicas (eventos da linha de varredura).
    # As coordenadas Y das peças vêm do rectpack (origem no canto inferior esquerdo),
    # mas nosso plano de corte inverte para o canto superior esquerdo. Precisamos usar o sistema de coordenadas do rectpack aqui.
    y_coords = {0, chapa_altura}
    for p in pecas_alocadas:
        y_coords.add(p['y'])
        y_coords.add(p['y'] + p['altura'])
    
    sorted_y = sorted(list(y_coords))
    sobras_brutas = []

    # 2. Iterar sobre cada faixa horizontal (strip) criada pelas coordenadas Y.
    for i in range(len(sorted_y) - 1):
        y1, y2 = sorted_y[i], sorted_y[i+1]
        altura_faixa = y2 - y1

        if altura_faixa < 1e-5:
            continue

        # Encontra todas as peças que se sobrepõem a esta faixa horizontal.
        mid_y = y1 + altura_faixa / 2
        pecas_na_faixa = [p for p in pecas_alocadas if p['y'] <= mid_y < (p['y'] + p['altura'])]
        
        # 3. Ordena as peças na faixa pelo eixo X para encontrar os vãos.
        pecas_na_faixa.sort(key=lambda p: p['x'])

        # 4. Encontra os espaços vazios (vãos) no eixo X.
        ponteiro_x = 0
        for p in pecas_na_faixa:
            # Se houver um espaço entre o ponteiro e a peça atual, é uma sobra.
            if p['x'] > ponteiro_x:
                largura_sobra = p['x'] - ponteiro_x
                if largura_sobra > 1e-5: # Ignora vãos insignificantes
                    sobras_brutas.append({
                        'x': ponteiro_x, 'y': y1,
                        'largura': largura_sobra, 'altura': altura_faixa
                    })
            # Atualiza o ponteiro para o final da peça atual.
            ponteiro_x = max(ponteiro_x, p['x'] + p['largura'])
        
        # 5. Verifica se há um último vão após a última peça até a borda da chapa.
        if ponteiro_x < chapa_largura:
            largura_sobra = chapa_largura - ponteiro_x
            if largura_sobra > 1e-5:
                sobras_brutas.append({
                    'x': ponteiro_x, 'y': y1,
                    'largura': largura_sobra, 'altura': altura_faixa
                })

    # 6. Fundir os retângulos de sobra adjacentes para formar peças maiores.
    sobras_fundidas = _merge_scraps(sobras_brutas)

    # 7. Filtrar pelo tamanho mínimo e classificar
    sobras_finais = []
    for s in sobras_fundidas:
        if s['largura'] >= min_dim and s['altura'] >= min_dim:
            # --- CORREÇÃO: Converte a coordenada Y para a origem no topo ANTES de retornar ---
            s['y'] = chapa_altura - s['y'] - s['altura']
            
            # Define 'tipo_sobra' para fins de relatório (e.g., no PDF)
            s['tipo_sobra'] = 'aproveitavel' if min(s['largura'], s['altura']) >= 300 and max(s['largura'], s['altura']) >= 300 else 'nao_aproveitavel'
            
            # --- INÍCIO: ATRIBUIÇÃO DE PONTUAÇÃO DE REUTILIZAÇÃO (TASK 2) ---
            # Esta pontuação ajuda a priorizar sobras mais úteis na fase de re-nesting.
            score = 0
            area = s['largura'] * s['altura']
            
            # Prioriza sobras muito grandes (e.g., > 10% da área da chapa bruta)
            if area > (chapa_largura * chapa_altura * 0.1):
                score += 10
            
            # Prioriza sobras que se encaixam na definição de "aproveitável" (mais quadradas/retangulares)
            if s['tipo_sobra'] == 'aproveitavel':
                score += 5
            
            # Prioriza faixas longas e estreitas (como as de 2980x... mm)
            if (s['largura'] >= chapa_largura * 0.9 and s['altura'] >= 100) or \
               (s['altura'] >= chapa_altura * 0.9 and s['largura'] >= 100):
                score += 7 # Alta prioridade para faixas longas
            
            s['potential_reuse_score'] = score
            # --- FIM: ATRIBUIÇÃO DE PONTUAÇÃO DE REUTILIZAÇÃO ---
            sobras_finais.append(s)

    logging.debug(f"Finalizado 'encontrar_sobras' (Scanline). Encontradas {len(sobras_finais)} sobras válidas.")
    return sobras_finais

def orquestrar_planos_de_corte(chapa_largura, chapa_altura, pecas, offset, margin, espessura, peso_especifico_base=7.85, status_signal_emitter=None):
    """
    Função mestre que orquestra o processo de nesting, priorizando o uso de sobras aproveitáveis.
    Esta é a função que deve ser chamada pela UI.
    """
    logging.info(f"--- INICIANDO ORQUESTRAÇÃO DE NESTING (ESTRATÉGIA OTIMIZADA) PARA ESPESSURA {espessura}mm ---")

    # --- INÍCIO: NOVA ESTRATÉGIA DE EMPACOTAMENTO ÚNICO E OTIMIZADO ---
    # 1. Ordena todas as peças, da maior para a menor, para que o algoritmo de empacotamento
    #    posicione as peças mais difíceis (maiores) primeiro. Isso geralmente leva a um
    #    resultado global melhor, pois as peças menores podem preencher os espaços restantes.
    pecas_ordenadas = sorted(pecas, key=lambda p: p['largura'] * p['altura'], reverse=True)
    logging.info(f"Ordenando {len(pecas_ordenadas)} tipos de peças por área para otimização.")

    # 2. Define os "bins" (chapas) disponíveis para o empacotamento.
    #    Nesta estratégia simplificada, usamos apenas chapas novas. A reutilização de sobras
    #    pode ser reintroduzida em uma etapa posterior, se necessário, mas o foco aqui
    #    é maximizar o aproveitamento em chapas virgens.
    bins_disponiveis = [(chapa_largura, chapa_altura, margin)] * 200 # Fornece um grande número de chapas brutas

    # 3. Executa o cálculo de nesting em uma única passagem com todas as peças.
    if status_signal_emitter: status_signal_emitter.emit("Otimizando alocação de todas as peças...")
    resultado_otimizado = calcular_plano_de_corte_em_bins(
        pecas_ordenadas, 
        offset, 
        espessura, 
        bins_disponiveis, 
        peso_especifico_base, 
        status_signal_emitter
    )

    logging.info(f"--- ORQUESTRAÇÃO OTIMIZADA FINALIZADA ---")
    return resultado_otimizado
    # --- FIM: NOVA ESTRATÉGIA ---


def calcular_plano_de_corte_em_bins(pecas, offset, espessura, bins, peso_especifico_base=7.85, status_signal_emitter=None):
    """
    Calcula o plano de corte, incluindo uma análise detalhada de pesos e sucatas.
    """
    logging.info(f"Iniciando cálculo de corte para {len(pecas)} tipos de peças em {len(bins)} bins disponíveis.")

    # Função interna para calcular peso com base na área, espessura e densidade
    def _calc_peso(area_mm2):
        if espessura is None or espessura <= 0:
            return 0
        # Formula: (Área em m²) * espessura em m * densidade (kg/m³)
        # Densidade do aço: 7.85 g/cm³ = 7850 kg/m³
        # A fórmula (area_mm2 / 1_000_000) * espessura * 7.85 é uma simplificação que chega no mesmo resultado.
        return (area_mm2 / 1_000_000) * espessura * peso_especifico_base

    todos_algoritmos = [MaxRectsBssf, MaxRectsBaf, MaxRectsBlsf, SkylineBl, SkylineMwf, SkylineBlWm, SkylineMwfl]
    melhor_resultado = None
    menor_num_chapas = float('inf')

    # --- LÓGICA DE FUSÃO DE FORMAS (Triângulos, Trapézios) ---
    pecas_processadas = []
    triangulos = [p for p in pecas if p.get('forma') == 'right_triangle']
    trapezios = [p for p in pecas if p.get('forma') == 'trapezoid']
    outras_pecas = [p for p in pecas if p.get('forma') not in ['right_triangle', 'trapezoid']]

    from collections import defaultdict
    mapa_triangulos = defaultdict(list)
    for t in triangulos:
        mapa_triangulos[(t['largura'], t['altura'])].append(t)

    for dim, lista_triangulos in mapa_triangulos.items():
        total_qtd = sum(t['quantidade'] for t in lista_triangulos)
        num_pares, num_sozinhos = divmod(total_qtd, 2)
        if num_pares > 0:
            pecas_processadas.append({'forma': 'paired_triangle', 'largura': dim[0], 'altura': dim[1], 'quantidade': num_pares})
        if num_sozinhos > 0:
            pecas_processadas.append({'forma': 'right_triangle', 'largura': dim[0], 'altura': dim[1], 'quantidade': num_sozinhos})

    mapa_trapezios = defaultdict(list)
    for t in trapezios:
        mapa_trapezios[(t['largura'], t['altura'], t.get('small_base', 0))].append(t)

    for dim, lista_trapezios in mapa_trapezios.items():
        total_qtd = sum(t['quantidade'] for t in lista_trapezios)
        num_pares, num_sozinhos = divmod(total_qtd, 2)
        if num_pares > 0:
            pecas_processadas.append({
                'forma': 'paired_trapezoid', 'largura': dim[0] + dim[2], 'altura': dim[1], 'quantidade': num_pares,
                'orig_dims': {'large_base': dim[0], 'small_base': dim[2], 'height': dim[1]}
            })
        if num_sozinhos > 0:
            pecas_processadas.append(lista_trapezios[0])

    pecas_processadas.extend(outras_pecas)
    # --- FIM DA LÓGICA DE FUSÃO ---

    peca_unica_id = 1
    retangulos_para_alocar = []
    id_peca_map = {}
    area_total_pecas_sem_offset = 0

    for peca_proc in pecas_processadas:
        # As dimensões em 'peca_proc' já incluem o offset, vindo de main.py
        largura_com_offset = peca_proc['largura']
        altura_com_offset = peca_proc['altura']
        
        # Calcula a área sem offset para o cálculo da área do offset posteriormente
        largura_sem_offset = largura_com_offset - offset if largura_com_offset > offset else largura_com_offset
        altura_sem_offset = altura_com_offset - offset if altura_com_offset > offset else altura_com_offset
        
        for _ in range(peca_proc['quantidade']):
            rid = str(peca_unica_id)
            retangulos_para_alocar.append((largura_com_offset, altura_com_offset, rid))
            
            # Armazena dados originais e com offset para referência
            id_peca_map[rid] = {
                'largura_com_offset': largura_com_offset, 'altura_com_offset': altura_com_offset,
                'largura_sem_offset': largura_sem_offset, 'altura_sem_offset': altura_sem_offset,
                'furos': peca_proc.get('furos', []),
                'forma': peca_proc.get('forma', 'rectangle'),
                'diametro': peca_proc.get('diametro', 0),
                'orig_dims': peca_proc.get('orig_dims'),
                'dxf_path': peca_proc.get('dxf_path')
            }
            peca_unica_id += 1

    num_pecas = len(retangulos_para_alocar)
    if num_pecas > 500: algoritmos_para_testar = [MaxRectsBssf]
    elif num_pecas > 200: algoritmos_para_testar = [MaxRectsBssf, MaxRectsBaf, SkylineMwf]
    else: algoritmos_para_testar = todos_algoritmos

    for algo in algoritmos_para_testar:
        if status_signal_emitter: status_signal_emitter.emit(f"Testando algoritmo: {algo.__name__}...")
        packer = rectpack.newPacker(rotation=True, pack_algo=algo)
        for r in retangulos_para_alocar: packer.add_rect(r[0], r[1], rid=r[2])
        
        # --- INÍCIO DA CORREÇÃO ---
        # Adiciona todos os bins disponíveis (sobras e chapas novas) ao packer
        for b_width, b_height, b_margin in bins:
            nesting_width = b_width - (2 * b_margin)
            nesting_height = b_height - (2 * b_margin)
            if nesting_width > 0 and nesting_height > 0:
                packer.add_bin(nesting_width, nesting_height, bid=(b_width, b_height, b_margin))

        packer.pack()
        # --- FIM DA CORREÇÃO ---

        if len(packer) < menor_num_chapas:
            menor_num_chapas = len(packer)
            logging.info(f"Novo melhor resultado: {menor_num_chapas} chapas com {algo.__name__}.")
            
            planos_agrupados = {}
            area_total_utilizada_com_offset = 0
            perda_intersticial_total = 0
            area_total_pecas_sem_offset_real = 0

            for bin_node in packer:
                chapa_largura, chapa_altura, margin = bin_node.bid
                nesting_width = chapa_largura - (2 * margin)
                nesting_height = chapa_altura - (2 * margin)
                
                # --- INÍCIO: CORREÇÃO PARA O ERRO 'tuple' object has no attribute 'x' ---
                # Filtra apenas os retângulos que foram empacotados com sucesso,
                # que são os objetos que possuem o atributo 'x'. As peças que não
                # couberam são retornadas como tuplas e serão ignoradas.
                chapa_alocada = [r for r in bin_node if hasattr(r, 'x')]
                # --- FIM: CORREÇÃO ---

                assinatura = tuple(sorted([(r.x, r.y, r.width, r.height) for r in chapa_alocada]))
                assinatura = (chapa_largura, chapa_altura) + assinatura
                
                if assinatura not in planos_agrupados:
                    plano_de_corte, pecas_contagem, perda_intersticial_plano = [], {}, 0
                    
                    for r in chapa_alocada:
                        peca_info = id_peca_map[r.rid]
                        forma = peca_info.get('forma', 'rectangle')
                        
                        # Lógica de identificação da peça (tipo_key)
                        if forma == 'rectangle': tipo_key = f"R {peca_info['largura_sem_offset']:.0f}x{peca_info['altura_sem_offset']:.0f}"
                        elif forma == 'circle': tipo_key = f"C Ø{peca_info['diametro']:.0f}"
                        elif forma == 'paired_triangle': tipo_key = f"2T {peca_info['largura_sem_offset']:.0f}x{peca_info['altura_sem_offset']:.0f}"
                        elif forma == 'paired_trapezoid':
                            dims = peca_info['orig_dims']
                            tipo_key = f"2Z {dims['large_base']-offset:.0f}/{dims['small_base']-offset:.0f}x{dims['height']-offset:.0f}"
                        elif forma == 'dxf_shape': tipo_key = f"DXF: {os.path.basename(peca_info['dxf_path'])}"
                        else: tipo_key = f"{forma[0].upper()} {peca_info['largura_sem_offset']:.0f}x{peca_info['altura_sem_offset']:.0f}"
                        pecas_contagem[tipo_key] = pecas_contagem.get(tipo_key, 0) + 1

                        # Cálculo de perda intersticial para formas não retangulares
                        area_bounding_box = peca_info['largura_com_offset'] * peca_info['altura_com_offset']
                        if forma == 'circle':
                            area_real_peca = math.pi * (peca_info['diametro'] / 2)**2
                            perda_intersticial_plano += (area_bounding_box - area_real_peca)
                        else: # Aprox. para retângulos e outras formas
                            area_real_peca = peca_info['largura_sem_offset'] * peca_info['altura_sem_offset']
                        
                        # Lógica de furos e rotação
                        furos_trans = []
                        foi_rotacionada = r.width != peca_info['largura_com_offset']
                        if foi_rotacionada:
                            for furo in peca_info['furos']:
                                furos_trans.append({'diam': furo['diam'], 'x': furo['y'], 'y': peca_info['largura_com_offset'] - furo['x']})
                        else:
                            furos_trans = peca_info['furos']
                        
                        plano_de_corte.append({
                            # Adiciona a margem para obter a coordenada na chapa real
                            "x": r.x + margin, 
                            # O Y do rectpack é de baixo para cima. Convertendo para cima para baixo na chapa real: 
                            "y": chapa_altura - (r.y + margin) - r.height, 
                            "largura": r.width, "altura": r.height,
                            "tipo_key": tipo_key, "furos": furos_trans, "forma": forma, "rid": r.rid,
                            "diametro": peca_info['diametro'], "orig_dims": peca_info.get('orig_dims'), "dxf_path": peca_info['dxf_path']
                        })
                        
                    resumo_pecas = [{"tipo": t, "qtd": q} for t, q in pecas_contagem.items()]
                    pecas_para_geometria = [{'x': r.x, 'y': r.y, 'largura': r.width, 'altura': r.height} for r in chapa_alocada]
                    # Encontra sobras na área de nesting e depois ajusta suas coordenadas. Usa chapa_alocada.
                    sobras_na_area_nesting = encontrar_sobras(nesting_width, nesting_height, pecas_para_geometria) # Usa as dimensões do bin atual
                    for s in sobras_na_area_nesting:
                        s['x'] += margin
                        s['y'] += margin # A coordenada Y da sobra já foi invertida para a origem no topo
                    sobras_encontradas = sobras_na_area_nesting

                    planos_agrupados[assinatura] = {
                        "plano": plano_de_corte, "repeticoes": 1, 
                        "resumo_pecas": resumo_pecas, "sobras": sobras_encontradas,
                        "chapa_largura": chapa_largura, "chapa_altura": chapa_altura # Armazena as dimensões do bin
                    }
                else:
                    planos_agrupados[assinatura]["repeticoes"] += 1
                
                area_total_utilizada_com_offset += sum(r.width * r.height for r in chapa_alocada)
                perda_intersticial_total += perda_intersticial_plano
                area_total_pecas_sem_offset_real += sum(id_peca_map[r.rid]['largura_sem_offset'] * id_peca_map[r.rid]['altura_sem_offset'] for r in chapa_alocada)
            
            # Calcula a área total das chapas/sobras realmente usadas
            area_total_chapas = sum(plano['chapa_largura'] * plano['chapa_altura'] * plano['repeticoes'] for plano in planos_agrupados.values())

            planos_unicos = list(planos_agrupados.values())
            total_chapas = sum(plano['repeticoes'] for plano in planos_unicos) 
            
            # A área real das peças é a soma das áreas sem offset.
            area_real_pecas = sum(
                (id_peca_map[r['rid']]['largura_sem_offset'] * id_peca_map[r['rid']]['altura_sem_offset']) * plano['repeticoes'] 
                for plano in planos_unicos for r in plano['plano'])
            area_utilizada_real = area_real_pecas # A perda intersticial já está contida na diferença para o bounding box
            aproveitamento_geral = (area_utilizada_real / area_total_chapas) * 100 if area_total_chapas > 0 else 0

            # --- INÍCIO: CÁLCULO DETALHADO DE SUCATA E PESOS ---
            total_area_sobra_aproveitavel, total_area_sobra_sucata = 0, 0
            sobras_aproveitaveis_detalhado, sucatas_dimensionadas_detalhado = [], []

            for plano in planos_unicos:
                for sobra in plano.get('sobras', []):
                    area_sobra = sobra['largura'] * sobra['altura']
                    item = {'largura': sobra['largura'], 'altura': sobra['altura'], 'peso': _calc_peso(area_sobra), 'quantidade': plano['repeticoes']}
                    if sobra['tipo_sobra'] == 'aproveitavel':
                        total_area_sobra_aproveitavel += area_sobra * plano['repeticoes']
                        sobras_aproveitaveis_detalhado.append(item)
                    else:
                        total_area_sobra_sucata += area_sobra * plano['repeticoes']
                        sucatas_dimensionadas_detalhado.append(item)
            
            # Calcula a área do offset como a diferença entre a área ocupada com e sem offset
            area_offset_total = area_total_utilizada_com_offset - area_real_pecas
            
            # Área das "demais sucatas" (perda de processo) é o que sobra após subtrair tudo
            area_demais_sucatas = area_total_chapas - area_utilizada_real - total_area_sobra_aproveitavel - total_area_sobra_sucata - area_offset_total
            # Garante que não seja negativo devido a erros de arredondamento
            if area_demais_sucatas < 0:
                area_demais_sucatas = 0
            
            sucata_detalhada = {
                "peso_offset": _calc_peso(area_offset_total),
                "sobras_aproveitaveis": sobras_aproveitaveis_detalhado,
                "sucatas_dimensionadas": sucatas_dimensionadas_detalhado,
                "peso_demais_sucatas": _calc_peso(area_demais_sucatas) # Renomeado para "perda de processo" no PDF
            }
            # --- FIM: CÁLCULO DETALHADO ---

            melhor_resultado = {
                "planos_unicos": planos_unicos,
                "total_chapas": total_chapas,
                "aproveitamento_geral": f"{aproveitamento_geral:.2f}%",
                "color_map": {},
                "area_total_chapas": area_total_chapas,
                "area_utilizada_real": area_utilizada_real,
                "total_area_sobra_aproveitavel": total_area_sobra_aproveitavel,
                "total_area_sobra_sucata": total_area_sobra_sucata,
                "sucata_detalhada": sucata_detalhada # NOVO CAMPO
            }
            
    if melhor_resultado:
        logging.info(f"Cálculo finalizado. Melhor resultado: {melhor_resultado['total_chapas']} chapas, {melhor_resultado['aproveitamento_geral']} de aproveitamento.")
    else:
        logging.warning("Nenhum resultado de cálculo foi gerado.")
    return melhor_resultado