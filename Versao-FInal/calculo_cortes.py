import rectpack
import logging
from PyQt5.QtCore import QObject, pyqtSignal
from rectpack.maxrects import MaxRectsBssf, MaxRectsBaf, MaxRectsBlsf, MaxRectsBl
from rectpack.skyline import SkylineBl, SkylineBlWm, SkylineMwf, SkylineMwfl

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
    Este processo é repetido até que não haja mais fusões possíveis.
    """
    if not scraps:
        return []

    while True:
        merged_in_pass = False
        i = 0
        while i < len(scraps):
            j = i + 1
            while j < len(scraps):
                r1 = scraps[i]
                r2 = scraps[j]

                # Tenta fusão vertical
                if r1['x'] == r2['x'] and r1['largura'] == r2['largura']:
                    if abs((r1['y'] + r1['altura']) - r2['y']) < 1e-5: # r2 está abaixo de r1
                        r1['altura'] += r2['altura']
                        scraps.pop(j)
                        merged_in_pass = True
                        continue
                    elif abs((r2['y'] + r2['altura']) - r1['y']) < 1e-5: # r1 está abaixo de r2
                        r1['y'] = r2['y']
                        r1['altura'] += r2['altura']
                        scraps.pop(j)
                        merged_in_pass = True
                        continue

                # Tenta fusão horizontal
                if r1['y'] == r2['y'] and r1['altura'] == r2['altura']:
                    if abs((r1['x'] + r1['largura']) - r2['x']) < 1e-5: # r2 está à direita de r1
                        r1['largura'] += r2['largura']
                        scraps.pop(j)
                        merged_in_pass = True
                        continue
                j += 1
            i += 1
        if not merged_in_pass:
            break
    return scraps

def encontrar_sobras(chapa_largura, chapa_altura, pecas_alocadas, min_dim=50):
    """
    Encontra os maiores retângulos de sobra em uma chapa.
    Usa um packer reverso para "empacotar" espaços vazios.
    
    :param chapa_largura: Largura da chapa.
    :param chapa_altura: Altura da chapa.
    :param pecas_alocadas: Lista de retângulos (peças) já posicionados.
    :param min_dim: Dimensão mínima (largura ou altura) para uma sobra ser considerada.
    :return: Lista de dicionários representando as sobras encontradas.
    """
    logging.debug(f"Iniciando 'encontrar_sobras' para chapa {chapa_largura}x{chapa_altura} com {len(pecas_alocadas)} peças.")

    # --- INÍCIO: CORREÇÃO DEFINITIVA - ABORDAGEM DE SUBTRAÇÃO GEOMÉTRICA ---
    # A biblioteca não suporta 'dont_pack'. A abordagem correta é calcular
    # as sobras manualmente subtraindo as áreas das peças da área da chapa.

    # Começamos com uma lista de sobras que é a própria chapa.
    sobras = [{'x': 0, 'y': 0, 'largura': chapa_largura, 'altura': chapa_altura}]

    for peca in pecas_alocadas:
        novas_sobras = []
        # A cada peça, iteramos sobre a lista de sobras atual e a substituímos
        # pelos "cortes" resultantes.
        sobras_a_processar = sobras
        sobras = []
        for sobra in sobras_a_processar:
            # Calcula a área de intersecção
            inter_x = max(sobra['x'], peca['x'])
            inter_y = max(sobra['y'], peca['y'])
            inter_w = min(sobra['x'] + sobra['largura'], peca['x'] + peca['largura']) - inter_x
            inter_h = min(sobra['y'] + sobra['altura'], peca['y'] + peca['altura']) - inter_y

            # Se houver intersecção (w e h > 0), divide a sobra
            if inter_w > 0 and inter_h > 0:
                # Sobra acima da intersecção
                if sobra['y'] < inter_y:
                    sobras.append({'x': sobra['x'], 'y': sobra['y'], 'largura': sobra['largura'], 'altura': inter_y - sobra['y']})
                # Sobra abaixo da intersecção
                if sobra['y'] + sobra['altura'] > inter_y + inter_h:
                    sobras.append({'x': sobra['x'], 'y': inter_y + inter_h, 'largura': sobra['largura'], 'altura': (sobra['y'] + sobra['altura']) - (inter_y + inter_h)})
                # Sobra à esquerda da intersecção
                if sobra['x'] < inter_x:
                    sobras.append({'x': sobra['x'], 'y': inter_y, 'largura': inter_x - sobra['x'], 'altura': inter_h})
                # Sobra à direita da intersecção
                if sobra['x'] + sobra['largura'] > inter_x + inter_w:
                    sobras.append({'x': inter_x + inter_w, 'y': inter_y, 'largura': (sobra['x'] + sobra['largura']) - (inter_x + inter_w), 'altura': inter_h})
            else:
                # Se não há intersecção, a sobra continua intacta.
                sobras.append(sobra)

    # --- INÍCIO: ETAPA DE FUSÃO DAS SOBRAS ---
    sobras = _merge_scraps(sobras)

    # Filtra e classifica as sobras finais
    sobras_finais = []
    for s in sobras:
        if s['largura'] >= min_dim and s['altura'] >= min_dim:
            # --- CORREÇÃO: A regra deve incluir 300 e considerar rotação ---
            # Uma peça é aproveitável se ambas as dimensões forem >= 300.
            tipo_sobra = 'aproveitavel' if min(s['largura'], s['altura']) >= 300 and max(s['largura'], s['altura']) >= 300 else 'nao_aproveitavel'
            sobras_finais.append({**s,
                "tipo_sobra": tipo_sobra
            })
    logging.debug(f"Finalizado 'encontrar_sobras'. Encontradas {len(sobras_finais)} sobras válidas.")
    # --- FIM: CORREÇÃO ---
    return sobras_finais

def calcular_plano_de_corte(chapa_largura, chapa_altura, pecas, status_signal_emitter=None):
    """
    Calcula o plano de corte testando múltiplos algoritmos de empacotamento
    e selecionando o que utiliza o menor número de chapas.
    """
    # --- INÍCIO: MELHORIA NO ALGORITMO ---
    logging.info(f"Iniciando cálculo de corte para chapa {chapa_largura}x{chapa_altura}.")
    
    # --- INÍCIO: OTIMIZAÇÃO DE PERFORMANCE ---
    # A lista completa de algoritmos a serem testados.
    todos_algoritmos = [MaxRectsBssf, MaxRectsBaf, MaxRectsBlsf, SkylineBl, SkylineMwf]
    melhor_resultado = None
    menor_num_chapas = float('inf')
    
    # --- INÍCIO: LÓGICA DE FUSÃO DE TRIÂNGULOS ---
    pecas_processadas = []
    triangulos = [p for p in pecas if p.get('forma') == 'right_triangle']
    outras_pecas = [p for p in pecas if p.get('forma') != 'right_triangle']

    # Agrupa triângulos por dimensões
    from collections import defaultdict
    mapa_triangulos = defaultdict(list)
    for t in triangulos:
        mapa_triangulos[(t['largura'], t['altura'])].append(t)

    for dim, lista_triangulos in mapa_triangulos.items():
        total_qtd = sum(t['quantidade'] for t in lista_triangulos)
        num_pares = total_qtd // 2
        num_sozinhos = total_qtd % 2

        if num_pares > 0:
            pecas_processadas.append({
                'forma': 'paired_triangle',
                'largura': dim[0], 'altura': dim[1],
                'quantidade': num_pares
            })
        if num_sozinhos > 0:
            pecas_processadas.append({
                'forma': 'right_triangle',
                'largura': dim[0], 'altura': dim[1],
                'quantidade': num_sozinhos
            })
    
    pecas_processadas.extend(outras_pecas)
    # --- FIM: LÓGICA DE FUSÃO DE TRIÂNGULOS ---

    # --- MUDANÇA NA GERAÇÃO DE ID ---
    peca_unica_id = 1 # Inicializa o contador único
    retangulos_para_alocar = []
    id_peca_map = {}
    for i, peca_proc in enumerate(pecas_processadas): # Usa a lista processada
        for j in range(peca_proc['quantidade']):
            # Atribui o ID único e converte para string
            rid = str(peca_unica_id)
            
            retangulos_para_alocar.append((peca_proc['largura'], peca_proc['altura'], rid))
            # --- MUDANÇA: Armazena também os furos no mapa de IDs ---
            id_peca_map[rid] = {
                'largura': peca_proc['largura'], 'altura': peca_proc['altura'],
                'furos': peca_proc.get('furos', []),
                'forma': peca_proc.get('forma', 'rectangle'), # Armazena a forma original
                'diametro': peca_proc.get('diametro', 0) # Armazena o diâmetro original se for círculo
            }
            
            # Incrementa o contador para a próxima peça
            peca_unica_id += 1
    # --- FIM DA MUDANÇA ---

    # --- OTIMIZAÇÃO: Se houver muitas peças, use apenas o algoritmo mais rápido para evitar lentidão. ---
    if len(retangulos_para_alocar) > 200:
        algoritmos_para_testar = [MaxRectsBssf]
        logging.warning(f"Muitas peças ({len(retangulos_para_alocar)}). Usando modo rápido (apenas MaxRectsBssf).")
    else:
        algoritmos_para_testar = todos_algoritmos

    for algo in algoritmos_para_testar:
        if status_signal_emitter: status_signal_emitter.emit(f"Testando algoritmo: {algo.__name__}...")
        logging.debug(f"Testando algoritmo: {algo.__name__}")
        packer = rectpack.newPacker(rotation=True, pack_algo=algo)
        for r in retangulos_para_alocar:
            packer.add_rect(r[0], r[1], rid=r[2])
        
        total_de_retangulos = len(retangulos_para_alocar)
        packer.add_bin(chapa_largura, chapa_altura)
        packer.pack()
        
        # Adiciona chapas (bins) até que todas as peças sejam alocadas
        while sum(len(b) for b in packer) < total_de_retangulos:
            logging.debug(f"  Peças alocadas: {sum(len(b) for b in packer)}/{total_de_retangulos}. Adicionando nova chapa.")
            packer.add_bin(chapa_largura, chapa_altura)
            packer.pack()
        
        logging.debug(f"Algoritmo {algo.__name__} finalizou com {len(packer)} chapas.")
        # Se o resultado atual for melhor (menos chapas), armazena-o
        if len(packer) < menor_num_chapas:
            menor_num_chapas = len(packer)
            logging.info(f"Novo melhor resultado encontrado com {menor_num_chapas} chapas usando {algo.__name__}.")
            
            planos_agrupados = {}
            area_total_utilizada = 0
            area_chapa = chapa_largura * chapa_altura
            
            for chapa in packer:
                pecas_na_chapa = sorted([(r.x, r.y, r.width, r.height) for r in chapa])
                assinatura = tuple(pecas_na_chapa)
                
                if assinatura not in planos_agrupados:
                    area_chapa_utilizada = sum(r.width * r.height for r in chapa)
                    plano_de_corte = []
                    pecas_contagem = {}
                    
                    for retangulo in chapa:
                        tipo_key = f"{id_peca_map[retangulo.rid]['largura']}x{id_peca_map[retangulo.rid]['altura']}"
                        pecas_contagem[tipo_key] = pecas_contagem.get(tipo_key, 0) + 1
                        
                        # --- INÍCIO: LÓGICA DE FUROS E ROTAÇÃO ---
                        peca_original = id_peca_map[retangulo.rid]
                        furos_originais = peca_original.get('furos', [])
                        furos_transformados = []
                        
                        # Verifica se a peça foi rotacionada
                        foi_rotacionada = retangulo.width != peca_original['largura']

                        if foi_rotacionada:
                            # Transforma as coordenadas dos furos se a peça foi rotacionada 90 graus
                            for furo in furos_originais:
                                furos_transformados.append({
                                    'diam': furo['diam'],
                                    'x': furo['y'],
                                    'y': peca_original['largura'] - furo['x']
                                })
                        else:
                            furos_transformados = furos_originais
                        # --- FIM: LÓGICA DE FUROS E ROTAÇÃO ---

                        plano_de_corte.append({
                            "x": retangulo.x,
                            "y": chapa_altura - retangulo.y - retangulo.height,
                            "largura": retangulo.width, "altura": retangulo.height,
                            "tipo_key": f"{peca_original['largura']}x{peca_original['altura']}",
                            "furos": furos_transformados,
                            # Se for um par, a forma é 'paired_triangle', senão a forma original
                            "forma": peca_original.get('forma', 'rectangle'),
                            "diametro": peca_original['diametro'] # Passa o diâmetro original
                        })
                        
                    resumo_pecas = [{"tipo": tipo, "qtd": qtd} for tipo, qtd in pecas_contagem.items()]
                    
                    # --- INÍCIO: CÁLCULO DE SOBRAS ---
                    sobras_encontradas = encontrar_sobras(chapa_largura, chapa_altura, plano_de_corte)
                    # --- FIM: CÁLCULO DE SOBRAS ---

                    planos_agrupados[assinatura] = {
                        "plano": plano_de_corte,
                        "repeticoes": 1,
                        "resumo_pecas": resumo_pecas,
                        "sobras": sobras_encontradas # Adiciona as sobras ao resultado
                    }
                else:
                    planos_agrupados[assinatura]["repeticoes"] += 1
                
                area_total_utilizada += sum(r.width * r.height for r in chapa)
                
            planos_unicos = list(planos_agrupados.values())
            total_chapas = len(packer)
            area_total_chapas = total_chapas * chapa_largura * chapa_altura
            aproveitamento_geral = (area_total_utilizada / area_total_chapas) * 100 if area_total_chapas > 0 else 0
            
            melhor_resultado = {
                "planos_unicos": planos_unicos,
                "total_chapas": total_chapas,
                "aproveitamento_geral": f"{aproveitamento_geral:.2f}%",
                "color_map": {} # Placeholder, será preenchido na UI
            }
            
    if melhor_resultado:
        logging.info(f"Cálculo finalizado. Melhor resultado: {melhor_resultado['total_chapas']} chapas, {melhor_resultado['aproveitamento_geral']} de aproveitamento.")
    else:
        logging.warning("Nenhum resultado de cálculo foi gerado.")
    return melhor_resultado
    # --- FIM: MELHORIA NO ALGORITMO ---