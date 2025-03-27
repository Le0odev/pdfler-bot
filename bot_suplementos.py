import fitz  # PyMuPDF
import re
import os
import mysql.connector
from mysql.connector import Error
import time
from difflib import SequenceMatcher
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import shutil
import threading
import traceback
import PyPDF2

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def extrair_texto_pdf(caminho_pdf):
    try:
        texto_completo = ""
        with fitz.open(caminho_pdf) as pdf:
            for pagina in pdf:
                texto_completo += pagina.get_text("text") + "\n"
        return texto_completo
    except Exception as e:
        return f"Erro ao extrair texto do PDF: {e}"

def processar_dados(texto):
    produtos = []
    if isinstance(texto, str) and "Erro" in texto:
        return produtos  # Retorna vazio se houve erro na extração
    
    # Imprime o texto para debug
    print("Texto extraído do PDF:")
    print(texto[:1000])  # Primeiros 1000 caracteres para debug
    
    # Método específico para o formato do PDF exemplo - mantemos apenas este que funcionou bem
    print("Usando método específico para o formato do PDF...")
    
    # Padrão específico para a estrutura do PDF exemplo
    pattern = re.compile(r'(\d+)\s+(\d+|[A-Z]+\d+)\s+(.*?)\s+(\d+)\s+UN\s+R\$\s+(\d+[.,]\d+)\s+[-]+\s+R\$\s+(\d+[.,]\d+)', re.DOTALL)
    
    texto_processado = re.sub(r'\n+', ' ', texto)  # Substitui quebras de linha por espaços
    matches = pattern.findall(texto_processado)
    
    for match in matches:
        try:
            item_num = match[0]
            codigo = match[1]
            nome_produto_completo = match[2].strip()
            
            # Limpa o nome do produto removendo códigos e padronizando
            nome_produto = limpar_nome_produto(nome_produto_completo)
                
            quantidade = int(match[3])
            preco_unitario = float(match[5].replace(',', '.'))  # Preço líquido (após desconto)
            
            # Preserva informação original do produto para casos especiais
            info_original = {}
            
            # Verifica se é Creatina RedLion com informação de gramatura
            if "CREATINA REDLION" in nome_produto or "RED LION" in nome_produto_completo and "CREATINE" in nome_produto_completo:
                gram_match = re.search(r'(\d+)\s*(?:G\b|GR\b|GRS\b)', nome_produto_completo.upper())
                if gram_match:
                    info_original['gramatura'] = gram_match.group(1)
                    if info_original['gramatura'] == '150':
                        nome_produto = "CREATINA REDLION 150G"
                    elif info_original['gramatura'] == '300':
                        nome_produto = "CREATINA REDLION 300G"
            
            # Verifica se é Thermo Abdomen com informação de cápsulas
            if "THERMO ABDOMEN" in nome_produto or "THERMO ABDOMEM" in nome_produto_completo:
                caps_match = re.search(r'(\d+)\s*(?:CAPS|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)', nome_produto_completo.upper())
                if caps_match:
                    info_original['caps'] = caps_match.group(1)
                    if info_original['caps'] == '60':
                        nome_produto = "THERMO ABDOMEN 60 CAPS"
                    elif info_original['caps'] == '120':
                        nome_produto = "THERMO ABDOMEN 120 CAPS"
            
            produtos.append({
                "product_name": nome_produto,
                "estoque_quant": quantidade,
                "estoque_peso": 0.00,
                "valor": preco_unitario,
                "is_bulk": 0,
                "info_original": info_original
            })
            print(f"Produto extraído: {nome_produto} | Qtd: {quantidade} | Preço: {preco_unitario}")
        except Exception as e:
            print(f"Erro ao processar item: {e}")
    
    print(f"Total de produtos extraídos: {len(produtos)}")
    return produtos

def limpar_nome_produto(nome):
    """
    Extrai apenas o nome principal do produto, ignorando sabores, cores e outras variações
    """
    # Remove códigos numéricos e alfanuméricos
    nome = re.sub(r'^(\d+\s+)+', '', nome)
    nome = re.sub(r'\b[A-Z0-9]+\d+\b', '', nome)
    
    # Normaliza para maiúsculas
    nome = nome.upper()
    
    # Casos especiais que precisamos tratar primeiro
    
    # Caso especial para THERMO ABDOMEN
    caps_thermo = None
    if 'THERMO ABDOMEN' in nome or 'THERMO ABDOMEM' in nome:
        caps_match = re.search(r'(\d+)\s*(?:CAPS|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)', nome)
        if caps_match:
            caps_thermo = caps_match.group(1)
            if caps_thermo == '60':
                return 'THERMO ABDOMEN 60 CAPS'
            elif caps_thermo == '120':
                return 'THERMO ABDOMEN 120 CAPS'
        return 'THERMO ABDOMEN'
    
    # Caso especial para CREATINA RED LION
    if 'CREATINA' in nome and ('RED LION' in nome or 'REDLION' in nome):
        gram_match = re.search(r'(\d+)\s*(?:G\b|GR\b|GRS\b)', nome)
        if gram_match:
            gramatura = gram_match.group(1)
            if gramatura == '150':
                return 'CREATINA REDLION 150G'
            elif gramatura == '300':
                return 'CREATINA REDLION 300G'
        return 'CREATINA REDLION'
    
    # Detecção específica para produtos New Millen
    if 'NEW MILLEN' in nome and ('C4' in nome or 'CAFFEINE' in nome or 'CAFFEINE FREE' in nome or 'CAFFEINE FRE' in nome):
        if 'BETA PUMP' in nome or 'BETA' in nome:
            return 'C4 BETA PUMP'
        if 'CAFFEINE' in nome or 'CAFFEINE FREE' in nome or 'CAFFEINE FRE' in nome:
            return 'C4 CAFFEINE FREE'
    
    if 'NEW MILLEN' in nome and 'PRE TREINO' in nome:
        if 'BETA PUMP' in nome or 'BETA' in nome:
            return 'C4 BETA PUMP'
    
    # Informação sobre cápsulas (para uso posterior na comparação)
    info_capsulas = None
    caps_match = re.search(r'(\d+)\s*(?:CAPS|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)', nome)
    if caps_match:
        info_capsulas = caps_match.group(0)
    
    # Primeiro, verifica se é um produto completo conhecido
    produtos_conhecidos = {
        'NATUROVOS ALBUMINA': ['NATUROVOS ALBUMINA', 'ALBUMINA NATUROVOS', 'ALBUMINA SABOR', 'NATUROVOS - ALBUMINA'],
        'C4 BETA PUMP': ['C4 BETA PUMP', 'BETA PUMP C4', 'NEW MILLEN - C4 BETA PUMP', 'NEW MILLEN C4 BETA', 'NEW MILLEN PRE TREINO BETA'],
        'C4 CAFFEINE FREE': ['C4 CAFFEINE FREE', 'C4 CAFFEINE FRE', 'NEW MILLEN - C4 CAFFEINE', 'NEW MILLEN CAFFEINE', 'NEW MILLEN PRE TREINO CAFFEINE'],
        'CREATINA REDLION 300G': ['RED LION CREATINE 300G', 'CREATINA RED LION 300G', 'RED LION SUPLEMENTOS - CREATINE 300G', 'CREATINE RED LION 300G'],
        'CREATINA REDLION 150G': ['RED LION CREATINE 150G', 'CREATINA RED LION 150G', 'RED LION SUPLEMENTOS - CREATINE 150G', 'CREATINE RED LION 150G'],
        'CREATINA REDLION': ['RED LION CREATINE', 'CREATINA RED LION', 'RED LION SUPLEMENTOS - CREATINE', 'CREATINE RED LION'],
        'THERMO ABDOMEN 60 CAPS': ['THERMO ABDOMEN 60 CAPS', 'BODYACTION - THERMO ABDOMEM 60 CAPS', 'THERMO ABDOMEM 60 CAPS'],
        'THERMO ABDOMEN 120 CAPS': ['THERMO ABDOMEN 120 CAPS', 'BODYACTION - THERMO ABDOMEM 120 CAPS', 'THERMO ABDOMEM 120 CAPS'],
        'THERMO ABDOMEN': ['THERMO ABDOMEN', 'BODYACTION - THERMO ABDOMEM', 'THERMO ABDOMEM'],
        'CREATINA BLACKSKULL': ['BLACKSKULL CREATINE', 'CREATINA BLACKSKULL', 'BLACK SKULL - CREATINE', 'CREATINE BLACK SKULL'],
        'GLUTAMINA REDLION': ['RED LION GLUTAMINE', 'GLUTAMINA RED LION', 'RED LION SUPLEMENTOS - GLUTAMINE', 'GLUTAMINE RED LION'],
        'BETA ALANINE REDLION': ['RED LION BETA ALANINE', 'BETA ALANINA RED LION', 'RED LION SUPLEMENTOS - BETA ALANINE', 'BETA ALANINE RED LION'],
        'ALFAJOR DR PEANUT': ['DR PEANUT ALFAJOR', 'ALFAJOR DR PEANUT', 'DR. PEANUT - ALFAJOR', 'DR. PEANULT - ALFAJOR'],
        'PASTA DE AMENDOIM DR PEANUT': ['DR PEANUT PASTA DE AMENDOIM', 'PASTA DE AMENDOIM DR PEANUT', 'DR. PEANUT - PASTA'],
        'MULTIVITAMINICO REDLION': ['RED LION MULTIVITAMINICO', 'MULTIVITAMINICO RED LION', 'RED LION SUPLEMENTOS - MULTIVITAMINICO'],
        'INSANE CLOWN DEMONS LAB': ['DEMONS LAB INSANE CLOWN', 'INSANE CLOWN DEMONS', 'DEMONS LAB - INSANE CLOWN'],
        'INSANE ORIGINAL DEMONS LAB': ['DEMONS LAB INSANE ORIGINAL', 'INSANE ORIGINAL DEMONS', 'DEMONS LAB - INSANE ORIGINAL'],
        'ECTOPLASMA DEMONS LAB': ['DEMONS LAB ECTOPLASMA', 'ECTOPLASMA DEMONS', 'DEMONS LAB - ECTOPLASMA']
    }
    
    # Verifica se o nome do produto corresponde a algum dos produtos conhecidos
    for produto_padrao, variantes in produtos_conhecidos.items():
        for variante in variantes:
            if variante in nome:
                # Se encontrou informação de cápsulas, a adiciona ao produto padronizado
                if info_capsulas and 'CAPS' not in produto_padrao:
                    return f"{produto_padrao} {info_capsulas}"
                return produto_padrao
    
    # Se não encontrou nas variantes conhecidas, tenta extrair a marca e o tipo de produto
    marcas = {
        'REDLION': ['RED LION SUPLEMENTOS', 'RED LION', 'REDLION'],
        'BLACKSKULL': ['BLACK SKULL', 'BLACKSKULL', 'BLACK'],
        'NATUROVOS': ['NATUROVOS'],
        'DEMONS LAB': ['DEMONS LAB', 'DEMONS', 'DEMON'],
        'NEW MILLEN': ['NEW MILLEN'],
        'INTEGRALMEDICA': ['INTEGRAL MEDICA', 'INTEGRALMEDICA'],
        'BODYACTION': ['BODYACTION', 'BODY ACTION'],
        'DR PEANUT': ['DR. PEANUT', 'DR. PEANULT', 'DR PEANUT']
    }
    
    tipos_produto = {
        'CREATINA': ['CREATINE', 'CREATINA'],
        'GLUTAMINA': ['GLUTAMINE', 'GLUTAMINA'],
        'ALBUMINA': ['ALBUMIN', 'ALBUMINA'],
        'BETA ALANINA': ['BETA ALANINE', 'BETA ALANINA'],
        'PRE TREINO': ['PRE WORKOUT', 'PRE-WORKOUT', 'PRE TREINO'],
        'THERMO ABDOMEN': ['THERMO ABDOMEM', 'THERMO ABDOMEN'],
        'WHEY PROTEIN': ['WHEY', 'PROTEIN', 'PROTEINA', 'WHEY PROTEIN'],
        'MULTIVITAMINICO': ['MULTIVITAMINICO', 'MULTIVITAMIN'],
        'C4': ['C4'],
        'BETA PUMP': ['BETA PUMP'],
        'INSANE': ['INSANE', 'INSANE CLOWN', 'INSANE ORIGINAL'],
        'ALFAJOR': ['ALFAJOR'],
        'PASTA DE AMENDOIM': ['PASTA DE AMENDOIM']
    }
    
    marca_encontrada = None
    for marca, variantes in marcas.items():
        for variante in variantes:
            if variante in nome:
                marca_encontrada = marca
                break
        if marca_encontrada:
            break
    
    tipo_encontrado = None
    for tipo, variantes in tipos_produto.items():
        for variante in variantes:
            if variante in nome:
                tipo_encontrado = tipo
                break
        if tipo_encontrado:
            break
            
    # Regras específicas para NEW MILLEN
    if marca_encontrada == 'NEW MILLEN' and tipo_encontrado == 'PRE TREINO':
        if 'BETA' in nome or 'PUMP' in nome:
            return 'C4 BETA PUMP'
        if 'CAFFEINE' in nome:
            return 'C4 CAFFEINE FREE'
    
    # Extrai informação sobre gramatura
    gram_match = re.search(r'(\d+)\s*(?:G\b|GR\b|GRS\b|ML\b|L\b)', nome)
    gramatura = None
    if gram_match:
        gramatura = gram_match.group(1)
    
    # Constrói o nome padrão do produto
    if marca_encontrada and tipo_encontrado:
        if tipo_encontrado == 'C4' and 'BETA PUMP' in nome:
            resultado = 'C4 BETA PUMP'
        elif tipo_encontrado == 'C4' and ('CAFFEINE' in nome or 'CAFFEINE FREE' in nome):
            resultado = 'C4 CAFFEINE FREE'
        else:
            resultado = f"{tipo_encontrado} {marca_encontrada}"
            
            # Adiciona gramatura para produtos específicos
            if tipo_encontrado == 'CREATINA' and marca_encontrada == 'REDLION' and gramatura:
                resultado = f"{tipo_encontrado} {marca_encontrada} {gramatura}G"
        
        # Adiciona informação de cápsulas se existir
        if info_capsulas:
            resultado += f" {info_capsulas}"
            
        return resultado
    
    # Caso não consiga padronizar, remove informações de sabor
    sabores = [
        'MORANGO', 'CHOCOLATE', 'NATURAL', 'BAUNILHA', 'BANANA', 
        'LARANJA', 'AMARELA', 'ROXA', 'LIMÃO', 'GRAPE', 'TUTTI', 
        'FRUTTI', 'FRUIT', 'PUNCH', 'COCO', 'BRIGADEIRO', 'AVELA',
        'SABOR'
    ]
    
    nome_processado = nome
    for sabor in sabores:
        nome_processado = re.sub(r'\b' + sabor + r'\b', '', nome_processado)
    
    # Preserve informações de cápsulas
    if info_capsulas:
        # Remove informações de gramatura que não são cápsulas
        nome_processado = re.sub(r'\d+\s*(?:G|GR|GRS|ML|L)', '', nome_processado)
        nome_processado = re.sub(r'\d+\s*UN', '', nome_processado)
    else:
        # Remove todas as informações de gramatura, caps, etc.
        nome_processado = re.sub(r'\d+\s*(?:G|GR|GRS|ML|L|CAPS?|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)', '', nome_processado)
        nome_processado = re.sub(r'\d+\s*UN', '', nome_processado)
    
    # Remove palavras comuns que não ajudam
    palavras_ignorar = ['DE', 'DA', 'DO', 'COM', 'SEM', 'SABOR', 'DISPLAY', 'WORKOUT', 'F.', 'F ']
    for palavra in palavras_ignorar:
        nome_processado = re.sub(r'\b' + palavra + r'\b', '', nome_processado)
    
    # Remove espaços extras e caracteres especiais
    nome_processado = re.sub(r'[-]', ' ', nome_processado)
    nome_processado = re.sub(r'\s+', ' ', nome_processado).strip()
    
    # Remove "SUPLEMENTOS" quando acompanha um nome de marca
    nome_processado = nome_processado.replace('SUPLEMENTOS', '')
    
    # Adiciona informação de cápsulas no final se existir
    if info_capsulas and info_capsulas not in nome_processado:
        nome_processado += f" {info_capsulas}"
    
    return nome_processado.strip()

def obter_produtos_do_banco():
    produtos = []
    try:
        conexao = mysql.connector.connect(
            host="autorack.proxy.rlwy.net",
            user="root",
            password="AGWseadASVhFzAaAlxmLBoYBzgvBQhVT",
            database="railway",
            port=16717
        )
        if conexao.is_connected():
            cursor = conexao.cursor(dictionary=True)
            cursor.execute("SELECT id, product_name, estoque_quant, estoque_peso, is_bulk FROM produto")
            produtos = cursor.fetchall()
            for produto in produtos:
                produto['estoque_quant'] = float(produto['estoque_quant']) if produto['estoque_quant'] is not None else 0
                produto['estoque_peso'] = float(produto['estoque_peso']) if produto['estoque_peso'] is not None else 0.0
    except Error as e:
        print(f"Erro ao conectar ao MySQL: {e}")
    finally:
        if 'conexao' in locals() and conexao.is_connected():
            cursor.close()
            conexao.close()
    return produtos

def normalizar_texto(texto):
    import unicodedata
    import re
    texto = texto.lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = re.sub(r'[^\w\s]', ' ', texto)
    
    # Padroniza os dados específicos
    substituicoes = {
        'org': 'organico', 
        'ml': '', 
        'g': '', 
        'kg': '', 
        'grs': '', 
        'gr': '', 
        'lt': '', 
        'l': '',
        'pct': '', 
        'un': '', 
        'cx': '', 
        'c/': 'com', 
        's/': 'sem', 
        'caps': '',
        'capsulas': '',
        'cap': '',
        'comp': '',
        'suplementos': '',
        'suplemento': '',
        'nutricao': '',
        'nutri': '',
        'pre workout': 'pre',
        'pre treino': 'pre',
        'pos treino': 'pos',
        '100%': '',
        'display': '',
        'workout': ''
    }
    
    for abrev, completo in substituicoes.items():
        texto = re.sub(r'\b' + abrev + r'\b', completo, texto)
    
    # Remover palavras comuns que não ajudam na comparação
    palavras_ignorar = ['de', 'da', 'do', 'em', 'a', 'o', 'e', 'com', 'sem', 'para', 'sabor', 'suplementos']
    palavras = texto.split()
    palavras_filtradas = [p for p in palavras if p not in palavras_ignorar]
    
    # Remover números isolados
    palavras_sem_numeros = [p for p in palavras_filtradas if not re.match(r'^\d+$', p)]
    
    # Preservar palavras-chave importantes mesmo após filtragem
    palavras_importantes = ['whey', 'creatina', 'albumina', 'glutamina', 'pre', 'pos', 'protein', 
                           'c4', 'beta', 'pump', 'hardcore', 'abdomen', 'redlion', 'blackskull', 'naturovos']
    
    # Readicionar palavras importantes caso tenham sido removidas
    texto_normalizado = ' '.join(palavras_sem_numeros)
    for palavra in palavras_importantes:
        if palavra in texto.lower() and palavra not in texto_normalizado:
            texto_normalizado += ' ' + palavra
    
    # Remover espaços extras
    texto_normalizado = re.sub(r'\s+', ' ', texto_normalizado).strip()
    
    return texto_normalizado

def extrair_informacoes_produto(nome):
    """Extrai informações relevantes do nome do produto"""
    info = {}
    nome_lower = nome.lower()
    
    # Extrai o tipo de produto
    tipos_produto = {
        'creatina': ['creatina', 'creatine'],
        'glutamina': ['glutamina', 'glutamine'],
        'albumina': ['albumina'],
        'whey': ['whey', 'protein'],
        'pre_treino': ['pre', 'pre workout', 'pre treino'],
        'c4': ['c4'],
        'beta_alanina': ['beta alanine'],
        'multivitaminico': ['multivitaminico'],
        'thermo': ['thermo'],
        'insane': ['insane']
    }
    
    for tipo, keywords in tipos_produto.items():
        for keyword in keywords:
            if keyword in nome_lower:
                info['tipo'] = tipo
                break
        if 'tipo' in info:
            break
    
    # Extrai o peso/volume
    peso_match = re.search(r'(\d+)\s*(?:g|gr|grs|mg|ml|l|caps)', nome_lower)
    if peso_match:
        info['peso'] = peso_match.group(1)
    
    # Extrai a marca
    marcas = {
        'redlion': ['redlion', 'red lion'],
        'blackskull': ['blackskull', 'black skull'],
        'naturovos': ['naturovos'],
        'demons': ['demons', 'demons lab'],
        'integralmedica': ['integralmedica', 'integral medica'],
        'bodyaction': ['bodyaction'],
        'new millen': ['new millen']
    }
    
    for marca, keywords in marcas.items():
        for keyword in keywords:
            if keyword in nome_lower:
                info['marca'] = marca
                break
        if 'marca' in info:
            break
    
    # Extrai o sabor se presente
    sabores = ['morango', 'chocolate', 'natural', 'laranja', 'amarela', 'roxa', 'grape', 'tutti', 'frutti']
    for sabor in sabores:
        if sabor in nome_lower:
            info['sabor'] = sabor
            break
    
    return info

def calcular_similaridade_produtos(nome1, nome2):
    """
    Calcula a similaridade entre produtos, buscando uma correspondência direta
    entre os tipos principais de produtos
    """
    # Padroniza os nomes dos produtos para comparação
    nome1_limpo = limpar_nome_produto(nome1)
    nome2_limpo = limpar_nome_produto(nome2)
    
    # Se os nomes são idênticos após limpeza, é uma correspondência perfeita
    if nome1_limpo == nome2_limpo:
        return 1.0
    
    # Verifica se um é uma variante ou substring do outro
    if nome1_limpo in nome2_limpo or nome2_limpo in nome1_limpo:
        # Quanto mais próximos em tamanho, mais similar
        razao_tamanho = min(len(nome1_limpo), len(nome2_limpo)) / max(len(nome1_limpo), len(nome2_limpo))
        return 0.8 + (razao_tamanho * 0.15)  # Entre 0.8 e 0.95
    
    # Produtos específicos que devem ser considerados o mesmo
    pares_equivalentes = [
        ('CREATINA BLACKSKULL', 'CREATINA REDLION'),
        ('CREATINA INTEGRALMEDICA', 'CREATINA REDLION'),
        ('CREATINA BODYACTION', 'CREATINA REDLION'),
        ('C4 BETA PUMP', 'PRE TREINO'),
        ('C4 CAFFEINE FREE', 'PRE TREINO'),
        ('ALBUMINA', 'NATUROVOS ALBUMINA')
    ]
    
    for par in pares_equivalentes:
        if (nome1_limpo == par[0] and nome2_limpo == par[1]) or (nome1_limpo == par[1] and nome2_limpo == par[0]):
            return 0.85  # Alta similaridade para produtos equivalentes
    
    # Verifica palavras-chave críticas em comum
    palavras_criticas = ['CREATINA', 'ALBUMINA', 'C4', 'PRE TREINO', 'BETA PUMP', 'GLUTAMINA', 'INSANE']
    
    for palavra in palavras_criticas:
        if palavra in nome1_limpo and palavra in nome2_limpo:
            return 0.75  # Boa similaridade para produtos com mesma palavra-chave crítica
    
    # Se não encontrou similaridade significativa
    return 0.0

def encontrar_produto_correspondente(nome_produto, produtos_banco):
    """
    Encontra o produto correspondente seguindo um fluxo de decisão estruturado
    que verifica múltiplas propriedades em sequência
    """
    print(f"\n========= BUSCANDO CORRESPONDÊNCIA =========")
    print(f"Produto: {nome_produto}")
    
    # Verificações para casos especiais
    
    # Caso especial: THERMO ABDOMEN
    if "THERMO ABDOMEN" in nome_produto.upper() or "THERMO ABDOMEM" in nome_produto.upper():
        caps_match = re.search(r'(\d+)\s*(?:CAPS|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)', nome_produto.upper())
        if caps_match:
            caps_val = caps_match.group(1)
            print(f"Verificando THERMO ABDOMEN com {caps_val} cápsulas")
            for produto in produtos_banco:
                if "THERMO ABDOMEN" in produto['product_name'].upper() or "THERMO ABDOMEM" in produto['product_name'].upper():
                    produto_caps_match = re.search(r'(\d+)\s*(?:CAPS|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)', produto['product_name'].upper())
                    if produto_caps_match and produto_caps_match.group(1) == caps_val:
                        print(f"✓✓ CORRESPONDÊNCIA EXATA PARA THERMO ABDOMEN {caps_val} CAPS com {produto['product_name']}")
                        return produto
            
            # Se não encontrou correspondência exata com o número de cápsulas, busca qualquer Thermo Abdomen
            for produto in produtos_banco:
                if "THERMO ABDOMEN" in produto['product_name'].upper() or "THERMO ABDOMEM" in produto['product_name'].upper():
                    print(f"✓ CORRESPONDÊNCIA APROXIMADA PARA THERMO ABDOMEN com {produto['product_name']}")
                    return produto
    
    # Caso especial: CREATINA REDLION
    if "CREATINA REDLION" in nome_produto.upper() or ("RED LION" in nome_produto.upper() and "CREATINE" in nome_produto.upper()):
        gram_match = re.search(r'(\d+)\s*(?:G\b|GR\b|GRS\b)', nome_produto.upper())
        if gram_match:
            gram_val = gram_match.group(1)
            print(f"Verificando CREATINA REDLION com {gram_val}g")
            for produto in produtos_banco:
                if ("CREATINA REDLION" in produto['product_name'].upper() or 
                    ("RED LION" in produto['product_name'].upper() and "CREATINE" in produto['product_name'].upper())):
                    produto_gram_match = re.search(r'(\d+)\s*(?:G\b|GR\b|GRS\b)', produto['product_name'].upper())
                    if produto_gram_match and produto_gram_match.group(1) == gram_val:
                        print(f"✓✓ CORRESPONDÊNCIA EXATA PARA CREATINA REDLION {gram_val}G com {produto['product_name']}")
                        return produto
            
            # Se não encontrou correspondência exata com a gramatura, usa 300g como padrão
            padrao = "300"  # Gramatura padrão se não especificada
            for produto in produtos_banco:
                if ("CREATINA REDLION" in produto['product_name'].upper() or 
                    ("RED LION" in produto['product_name'].upper() and "CREATINE" in produto['product_name'].upper())):
                    produto_gram_match = re.search(r'(\d+)\s*(?:G\b|GR\b|GRS\b)', produto['product_name'].upper())
                    if produto_gram_match and produto_gram_match.group(1) == padrao:
                        print(f"✓ CORRESPONDÊNCIA PARA CREATINA REDLION (usando padrão {padrao}g) com {produto['product_name']}")
                        return produto
    
    # Correções específicas para produtos conhecidos
    # NEW MILLEN PRE TREINO -> C4 BETA PUMP ou C4 CAFFEINE FREE
    nome_corrigido = nome_produto
    if 'NEW MILLEN' in nome_produto.upper():
        if 'CAFFEINE' in nome_produto.upper() or 'CAFFEINE FREE' in nome_produto.upper() or 'CAFFEINE FRE' in nome_produto.upper():
            nome_corrigido = 'C4 CAFFEINE FREE'
            print(f"Corrigindo nome para: {nome_corrigido}")
        elif 'BETA' in nome_produto.upper() or 'PUMP' in nome_produto.upper() or 'PRE TREINO' in nome_produto.upper():
            nome_corrigido = 'C4 BETA PUMP'
            print(f"Corrigindo nome para: {nome_corrigido}")
    
    # ETAPA 1: PREPARAÇÃO - Extrair e normalizar informações relevantes
    info_produto = extrair_propriedades_produto(nome_corrigido)
    nome_limpo = info_produto['nome_limpo']
    print(f"Nome normalizado: {nome_limpo}")
    
    if 'marca' in info_produto:
        print(f"Marca identificada: {info_produto['marca']}")
    if 'tipo' in info_produto:
        print(f"Tipo identificado: {info_produto['tipo']}")
    if 'gramatura' in info_produto:
        print(f"Gramatura identificada: {info_produto['gramatura']}")
    if 'caps' in info_produto:
        print(f"Cápsulas identificadas: {info_produto['caps']}")
    
    # Verificações especiais para produtos difíceis
    produtos_especiais = {
        'NEW MILLEN CAFFEINE FREE': ['C4 CAFFEINE FREE', 'C4 CAFFEINE FREE 30 SERVINGS', 'C4 CAFFEINE FREE 60 SERVINGS'],
        'NEW MILLEN CAFFEINE FRE': ['C4 CAFFEINE FREE', 'C4 CAFFEINE FREE 30 SERVINGS', 'C4 CAFFEINE FREE 60 SERVINGS'],
        'NEW MILLEN C4': ['C4 BETA PUMP', 'C4 BETA PUMP 30 SERVINGS', 'C4 BETA PUMP 60 SERVINGS'],
        'PRE TREINO NEW MILLEN': ['C4 BETA PUMP', 'C4 BETA PUMP 30 SERVINGS', 'C4 BETA PUMP 60 SERVINGS'],
        'CREATINA RED LION 300G': ['CREATINA REDLION 300G', 'CREATINA REDLION 300GR'],
        'CREATINA RED LION 150G': ['CREATINA REDLION 150G', 'CREATINA REDLION 150GR'],
        'THERMO ABDOMEN 60 CAPS': ['THERMO ABDOMEN 60 CAPS', 'THERMO ABDOMEN 60CAPS'],
        'THERMO ABDOMEN 120 CAPS': ['THERMO ABDOMEN 120 CAPS', 'THERMO ABDOMEN 120CAPS']
    }
    
    for nome_especial, correspondencias in produtos_especiais.items():
        if nome_especial.upper() in nome_produto.upper():
            for correspondencia in correspondencias:
                for produto in produtos_banco:
                    if correspondencia.upper() in produto['product_name'].upper():
                        print(f"✓✓ CORRESPONDÊNCIA ESPECIAL com {produto['product_name']}")
                        return produto
    
    # ETAPA 2: CORRESPONDÊNCIA EXATA - Nome exatamente igual após normalização
    print("\n--- Verificando correspondência exata ---")
    for produto in produtos_banco:
        info_banco = extrair_propriedades_produto(produto['product_name'])
        
        if nome_limpo.lower() == info_banco['nome_limpo'].lower():
            print(f"✓✓ CORRESPONDÊNCIA EXATA com {produto['product_name']}")
            return produto
    
    # ETAPA 3: CORRESPONDÊNCIA POR IDENTIDADE FORTE - Mesmo tipo e marca
    print("\n--- Verificando identidade forte (tipo + marca) ---")
    if 'tipo' in info_produto and 'marca' in info_produto:
        for produto in produtos_banco:
            info_banco = extrair_propriedades_produto(produto['product_name'])
            
            if ('tipo' in info_banco and 'marca' in info_banco and 
                info_produto['tipo'] == info_banco['tipo'] and 
                info_produto['marca'] == info_banco['marca']):
                
                # Verifica cápsulas/comprimidos se disponível
                if ('caps' in info_produto and 'caps' in info_banco and 
                    info_produto['caps'] == info_banco['caps']):
                    print(f"✓✓ CORRESPONDÊNCIA FORTE (tipo+marca+caps) com {produto['product_name']}")
                    return produto
                    
                # Verifica gramatura se disponível  
                if ('gramatura' in info_produto and 'gramatura' in info_banco and 
                    info_produto['gramatura'] == info_banco['gramatura']):
                    print(f"✓✓ CORRESPONDÊNCIA FORTE (tipo+marca+gramatura) com {produto['product_name']}")
                    return produto
                
                print(f"✓ CORRESPONDÊNCIA (tipo+marca) com {produto['product_name']}")
                return produto
    
    # ETAPA 4: CORRESPONDÊNCIA POR TIPO E ESPECIFICAÇÕES TÉCNICAS
    print("\n--- Verificando tipo e especificações técnicas ---")
    if 'tipo' in info_produto:
        candidatos_mesmo_tipo = []
        
        for produto in produtos_banco:
            info_banco = extrair_propriedades_produto(produto['product_name'])
            
            if 'tipo' in info_banco and info_produto['tipo'] == info_banco['tipo']:
                score = 0.6  # Pontuação base por ter o mesmo tipo
                motivo = "tipo"
                
                # Adiciona pontos por cápsulas iguais
                if 'caps' in info_produto and 'caps' in info_banco and info_produto['caps'] == info_banco['caps']:
                    score += 0.3
                    motivo += "+caps"
                
                # Adiciona pontos por gramatura igual
                if 'gramatura' in info_produto and 'gramatura' in info_banco and info_produto['gramatura'] == info_banco['gramatura']:
                    score += 0.3
                    motivo += "+gramatura"
                
                candidatos_mesmo_tipo.append((produto, score, motivo))
        
        # Se encontrou candidatos do mesmo tipo, retorna o de maior pontuação
        if candidatos_mesmo_tipo:
            candidatos_mesmo_tipo.sort(key=lambda x: x[1], reverse=True)
            melhor_candidato = candidatos_mesmo_tipo[0]
            print(f"✓ CORRESPONDÊNCIA POR {melhor_candidato[2]} com {melhor_candidato[0]['product_name']} (score: {melhor_candidato[1]:.2f})")
            return melhor_candidato[0]
    
    # ETAPA 5: CORRESPONDÊNCIA POR SUBSTRING - Um nome contém o outro
    print("\n--- Verificando correspondência por substring ---")
    for produto in produtos_banco:
        info_banco = extrair_propriedades_produto(produto['product_name'])
        
        if (nome_limpo.lower() in info_banco['nome_limpo'].lower() or 
            info_banco['nome_limpo'].lower() in nome_limpo.lower()):
            
            print(f"✓ CORRESPONDÊNCIA POR SUBSTRING com {produto['product_name']}")
            return produto
    
    # ETAPA 6: CORRESPONDÊNCIA POR PALAVRAS-CHAVE CRÍTICAS
    print("\n--- Verificando palavras-chave críticas ---")
    candidatos_palavras_chave = []
    
    palavras_criticas = ['CREATINA', 'GLUTAMINA', 'ALBUMINA', 'WHEY PROTEIN', 'PRE TREINO', 
                         'C4', 'BETA PUMP', 'INSANE', 'ALFAJOR', 'MULTIVITAMINICO']
    
    # Detecta palavras-chave no produto do PDF
    palavras_produto = set()
    for palavra in palavras_criticas:
        if palavra in nome_limpo.upper():
            palavras_produto.add(palavra)
    
    if palavras_produto:
        print(f"Palavras-chave no produto: {', '.join(palavras_produto)}")
        
        for produto in produtos_banco:
            info_banco = extrair_propriedades_produto(produto['product_name'])
            nome_banco = info_banco['nome_limpo'].upper()
            
            # Conta quantas palavras-chave existem no produto do banco
            matches = 0
            palavras_encontradas = []
            for palavra in palavras_produto:
                if palavra in nome_banco:
                    matches += 1
                    palavras_encontradas.append(palavra)
            
            if matches > 0:
                score = matches / len(palavras_produto)
                candidatos_palavras_chave.append((produto, score, palavras_encontradas))
    
    if candidatos_palavras_chave:
        candidatos_palavras_chave.sort(key=lambda x: x[1], reverse=True)
        melhor_candidato = candidatos_palavras_chave[0]
        palavras = ", ".join(melhor_candidato[2])
        print(f"✓ CORRESPONDÊNCIA POR PALAVRAS-CHAVE ({palavras}) com {melhor_candidato[0]['product_name']} (score: {melhor_candidato[1]:.2f})")
        return melhor_candidato[0]
    
    # ETAPA 7: BUSCA POR SIMILARIDADE GLOBAL
    print("\n--- Calculando similaridade global ---")
    candidatos_similaridade = []
    
    for produto in produtos_banco:
        similaridade = calcular_similaridade_global(nome_corrigido, produto['product_name'])
        
        if similaridade > 0.65:  # Limiar de alta similaridade
            candidatos_similaridade.append((produto, similaridade))
    
    if candidatos_similaridade:
        candidatos_similaridade.sort(key=lambda x: x[1], reverse=True)
        melhor_candidato = candidatos_similaridade[0]
        print(f"✓ CORRESPONDÊNCIA POR SIMILARIDADE GLOBAL com {melhor_candidato[0]['product_name']} (score: {melhor_candidato[1]:.2f})")
        return melhor_candidato[0]
    
    print("✗ Nenhuma correspondência encontrada")
    return None

def extrair_propriedades_produto(nome_produto):
    """
    Extrai todas as propriedades relevantes de um produto:
    - Nome limpo (sem códigos, sabores, etc)
    - Tipo de produto (creatina, whey, etc)
    - Marca (quando identificável)
    - Gramatura
    - Número de cápsulas/comprimidos
    """
    resultado = {}
    nome = nome_produto.upper()
    
    # Remove códigos e normaliza espaços
    nome = re.sub(r'^(\d+\s+)+', '', nome)
    nome = re.sub(r'\b[A-Z0-9]+\d+\b', '', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    
    # Extrai informação sobre cápsulas
    caps_match = re.search(r'(\d+)\s*(?:CAPS|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)', nome)
    if caps_match:
        resultado['caps'] = caps_match.group(1)  # Apenas o número
    
    # Extrai informação sobre gramatura
    gram_match = re.search(r'(\d+)\s*(?:G\b|GR\b|GRS\b|ML\b|L\b)', nome)
    if gram_match:
        resultado['gramatura'] = gram_match.group(1)  # Apenas o número
    
    # Dicionário de tipos de produtos
    tipos_produto = {
        'CREATINA': ['CREATINA', 'CREATINE'],
        'GLUTAMINA': ['GLUTAMINA', 'GLUTAMINE'],
        'ALBUMINA': ['ALBUMINA', 'ALBUMIN'],
        'BETA ALANINA': ['BETA ALANINA', 'BETA ALANINE'],
        'WHEY PROTEIN': ['WHEY PROTEIN', 'WHEY', 'PROTEIN', 'PROTEINA'],
        'PRE TREINO': ['PRE TREINO', 'PRE WORKOUT', 'PRE-WORKOUT'],
        'C4 BETA PUMP': ['C4 BETA PUMP', 'BETA PUMP C4'],
        'C4 CAFFEINE FREE': ['C4 CAFFEINE FREE', 'C4 CAFFEINE FRE'],
        'THERMO ABDOMEN': ['THERMO ABDOMEN', 'THERMO ABDOMEM'],
        'MULTIVITAMINICO': ['MULTIVITAMINICO', 'MULTIVITAMIN'],
        'INSANE CLOWN': ['INSANE CLOWN'],
        'INSANE ORIGINAL': ['INSANE ORIGINAL'],
        'ALFAJOR': ['ALFAJOR'],
        'PASTA DE AMENDOIM': ['PASTA DE AMENDOIM']
    }
    
    # Dicionário de marcas
    marcas = {
        'REDLION': ['RED LION SUPLEMENTOS', 'RED LION', 'REDLION'],
        'BLACKSKULL': ['BLACK SKULL', 'BLACKSKULL'],
        'NATUROVOS': ['NATUROVOS'],
        'DEMONS LAB': ['DEMONS LAB', 'DEMONS', 'DEMON'],
        'NEW MILLEN': ['NEW MILLEN'],
        'INTEGRALMEDICA': ['INTEGRAL MEDICA', 'INTEGRALMEDICA'],
        'BODYACTION': ['BODYACTION', 'BODY ACTION'],
        'DR PEANUT': ['DR. PEANUT', 'DR. PEANULT', 'DR PEANUT']
    }
    
    # Identifica o tipo de produto
    for tipo, variantes in tipos_produto.items():
        for variante in variantes:
            if variante in nome:
                resultado['tipo'] = tipo
                break
        if 'tipo' in resultado:
            break
    
    # Identifica a marca
    for marca, variantes in marcas.items():
        for variante in variantes:
            if variante in nome:
                resultado['marca'] = marca
                break
        if 'marca' in resultado:
            break
    
    # Remove sabores para criar nome limpo
    sabores = [
        'MORANGO', 'CHOCOLATE', 'NATURAL', 'BAUNILHA', 'BANANA', 
        'LARANJA', 'AMARELA', 'ROXA', 'LIMÃO', 'GRAPE', 'TUTTI', 
        'FRUTTI', 'FRUIT', 'PUNCH', 'COCO', 'BRIGADEIRO', 'AVELA',
        'SABOR'
    ]
    
    nome_sem_sabor = nome
    for sabor in sabores:
        nome_sem_sabor = re.sub(r'\b' + sabor + r'\b', '', nome_sem_sabor)
    
    # Remove palavras comuns que não ajudam
    palavras_ignorar = ['DE', 'DA', 'DO', 'COM', 'SEM', 'DISPLAY', 'WORKOUT', 'F.', 'F ']
    for palavra in palavras_ignorar:
        nome_sem_sabor = re.sub(r'\b' + palavra + r'\b', '', nome_sem_sabor)
    
    # Remove "SUPLEMENTOS" quando acompanha um nome de marca
    nome_sem_sabor = nome_sem_sabor.replace('SUPLEMENTOS', '')
    
    # Remove gramaturas e cápsulas do nome limpo
    nome_sem_sabor = re.sub(r'\d+\s*(?:G|GR|GRS|ML|L|CAPS|CAPSULAS|CÁPSULAS|COMPRIMIDOS|COMP)\b', '', nome_sem_sabor)
    
    # Limpa espaços duplicados e caracteres especiais
    nome_sem_sabor = re.sub(r'[-]', ' ', nome_sem_sabor)
    nome_sem_sabor = re.sub(r'\s+', ' ', nome_sem_sabor).strip()
    
    resultado['nome_limpo'] = nome_sem_sabor
    
    return resultado

def calcular_similaridade_global(nome1, nome2):
    """
    Calcula a similaridade global entre dois produtos considerando
    múltiplos fatores de similaridade
    """
    # Extrai propriedades de ambos os produtos
    props1 = extrair_propriedades_produto(nome1)
    props2 = extrair_propriedades_produto(nome2)
    
    # Inicializa pontuação
    score = 0.0
    
    # 1. Similaridade de nome limpo (30%)
    nome1_limpo = props1['nome_limpo'].lower()
    nome2_limpo = props2['nome_limpo'].lower()
    
    # Usa SequenceMatcher para comparar nomes limpos
    seq_similarity = SequenceMatcher(None, nome1_limpo, nome2_limpo).ratio()
    score += seq_similarity * 0.3
    
    # 2. Correspondência de tipo (30%)
    if 'tipo' in props1 and 'tipo' in props2:
        if props1['tipo'] == props2['tipo']:
            score += 0.3
    
    # 3. Correspondência de marca (15%)
    if 'marca' in props1 and 'marca' in props2:
        if props1['marca'] == props2['marca']:
            score += 0.15
    
    # 4. Correspondência de gramatura (15%)
    if 'gramatura' in props1 and 'gramatura' in props2:
        if props1['gramatura'] == props2['gramatura']:
            score += 0.15
    
    # 5. Correspondência de cápsulas (10%)
    if 'caps' in props1 and 'caps' in props2:
        if props1['caps'] == props2['caps']:
            score += 0.1
    
    return score

def atualizar_estoque(produtos_pdf, app=None):
    try:
        print("Iniciando atualização de estoque...")
        if app:
            app.log("🔄 Iniciando atualização de estoque...")
        
        produtos_banco = obter_produtos_do_banco()
        if not produtos_banco:
            print("Erro: Nenhum produto obtido do banco de dados")
            if app:
                app.log("❌ Erro: Nenhum produto obtido do banco de dados")
            return 0
        
        print(f"Produtos obtidos do banco: {len(produtos_banco)}")
        if app:
            app.log(f"Produtos obtidos do banco: {len(produtos_banco)}")
        
        try:
            print("Tentando conectar ao banco de dados...")
            conexao = mysql.connector.connect(
                host="autorack.proxy.rlwy.net",
                user="root",
                password="AGWseadASVhFzAaAlxmLBoYBzgvBQhVT",
                database="railway",
                port=16717
            )
            print("Conexão estabelecida com sucesso!")
            
            if conexao.is_connected():
                cursor = conexao.cursor()
                produtos_atualizados = 0
                produtos_processados = set()
                
                print("Processando produtos do PDF...")
                for produto_pdf in produtos_pdf:
                    print(f"Buscando correspondência para: {produto_pdf['product_name']}")
                    produto_banco = encontrar_produto_correspondente(produto_pdf['product_name'], produtos_banco)
                    
                    if produto_banco and produto_banco['id'] not in produtos_processados:
                        estoque_atual_quant = produto_banco['estoque_quant']
                        estoque_atual_peso = produto_banco['estoque_peso']
                        
                        if app:
                            app.log(f"Produto: {produto_banco['product_name']} | Estoque atual: {estoque_atual_quant}/{estoque_atual_peso} | A adicionar: {produto_pdf['estoque_quant']}/{produto_pdf['estoque_peso']}")
                        
                        try:
                            if produto_pdf['is_bulk'] == 1:
                                query = "UPDATE produto SET estoque_peso = estoque_peso + %s WHERE id = %s"
                                cursor.execute(query, (produto_pdf['estoque_peso'], produto_banco['id']))
                                novo_estoque = estoque_atual_peso + produto_pdf['estoque_peso']
                                print(f"Executou query para atualizar peso: {query} com valores {produto_pdf['estoque_peso'], produto_banco['id']}")
                            else:
                                query = "UPDATE produto SET estoque_quant = estoque_quant + %s WHERE id = %s"
                                cursor.execute(query, (produto_pdf['estoque_quant'], produto_banco['id']))
                                novo_estoque = estoque_atual_quant + produto_pdf['estoque_quant']
                                print(f"Executou query para atualizar quantidade: {query} com valores {produto_pdf['estoque_quant'], produto_banco['id']}")
                            
                            produtos_atualizados += 1
                            produtos_processados.add(produto_banco['id'])
                            
                            if app:
                                app.log(f"Atualizado: {produto_banco['product_name']} | Novo estoque: {novo_estoque}")
                        except Exception as query_error:
                            print(f"Erro ao executar query de atualização: {query_error}")
                            if app:
                                app.log(f"❌ Erro ao atualizar produto {produto_banco['product_name']}: {query_error}")
                    else:
                        print(f"Não encontrou correspondência para: {produto_pdf['product_name']}")
                        if app:
                            app.log(f"⚠️ Não encontrou correspondência para: {produto_pdf['product_name']}")
                
                print(f"Commit das alterações ({produtos_atualizados} produtos)...")
                conexao.commit()
                
                if app:
                    app.log(f"✅ {produtos_atualizados} produtos atualizados com sucesso.")
                
                return produtos_atualizados
        except Exception as conexao_error:
            print(f"Erro ao conectar ou processar com o banco de dados: {conexao_error}")
            if app:
                app.log(f"❌ Erro na conexão com o banco de dados: {conexao_error}")
            return 0
    except Error as e:
        print(f"Erro geral: {e}")
        if app:
            app.log(f"❌ Erro ao conectar ao MySQL: {e}")
        return 0
    finally:
        if 'conexao' in locals() and conexao.is_connected():
            cursor.close()
            conexao.close()
            print("Conexão ao MySQL fechada.")
            if app:
                app.log("Conexão ao MySQL fechada.")

class PDFBotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFler Bot - Suplementos")
        self.root.geometry("750x750")
        self.root.resizable(True, True)
        
        # Define um tamanho mínimo para a janela
        self.root.minsize(650, 500)
        
        self.pasta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pedidos")
        self.pasta_relatorios = os.path.join(os.path.dirname(os.path.abspath(__file__)), "relatorios")
        
        if not os.path.exists(self.pasta):
            os.makedirs(self.pasta)
        if not os.path.exists(self.pasta_relatorios):
            os.makedirs(self.pasta_relatorios)
        
        self.processados = set()
        self.ultimo_produtos = []  # Adicionando variável para armazenar os últimos produtos processados
        self.ultimo_relatorio = None
        self.resultados_comparacao = []  # Adicionando variável para armazenar os resultados
        
        self.main_frame = ctk.CTkFrame(root, corner_radius=10)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Frame superior para título e botões principais
        self.top_frame = ctk.CTkFrame(self.main_frame, corner_radius=5)
        self.top_frame.pack(pady=5, padx=10, fill="x")
        
        self.title_label = ctk.CTkLabel(self.top_frame, text="PDFler Bot - Suplementos", font=("Arial", 20, "bold"))
        self.title_label.pack(pady=10, side="left", padx=10)
        
        # Frame para botões de ação
        self.button_frame = ctk.CTkFrame(self.main_frame, corner_radius=5)
        self.button_frame.pack(pady=5, padx=10, fill="x")
        
        self.add_button = ctk.CTkButton(self.button_frame, text="📄 Adicionar PDF", command=self.adicionar_pdf, 
                                        font=("Arial", 12), width=150, height=35)
        self.add_button.pack(pady=5, padx=5, side="left")
        
        self.produtos_button = ctk.CTkButton(self.button_frame, text="🔍 Ver Produtos", command=self.ver_produtos, 
                                           font=("Arial", 12), width=150, height=35)
        self.produtos_button.pack(pady=5, padx=5, side="left")
        
        self.relatorio_button = ctk.CTkButton(self.button_frame, text="📊 Ver Relatório", command=self.ver_relatorio, 
                                            font=("Arial", 12), width=150, height=35)
        self.relatorio_button.pack(pady=5, padx=5, side="left")
        
        # Adicionar botão para atualizar estoque
        self.update_button = ctk.CTkButton(self.button_frame, text="📦 Atualizar Estoque", command=self.atualizar_estoque_manual, 
                                          font=("Arial", 12), width=170, height=35, fg_color="#28a745", hover_color="#218838")
        self.update_button.pack(pady=5, padx=5, side="left")
        
        # Frame para status e logs
        self.status_frame = ctk.CTkFrame(self.main_frame, corner_radius=5)
        self.status_frame.pack(pady=5, padx=10, fill="x")
        
        self.status_label = ctk.CTkLabel(self.status_frame, text="⏳ Aguardando PDF...", font=("Arial", 12))
        self.status_label.pack(pady=5, side="left", padx=10)
        
        self.open_folder_button = ctk.CTkButton(self.status_frame, text="📁 Abrir Pasta", command=self.abrir_pasta, 
                                               font=("Arial", 12), width=120, height=30)
        self.open_folder_button.pack(pady=5, padx=5, side="right")
        
        # Frame para o log
        self.log_frame = ctk.CTkFrame(self.main_frame, corner_radius=5)
        self.log_frame.pack(pady=5, padx=10, fill="both", expand=True)
        
        self.log_label = ctk.CTkLabel(self.log_frame, text="Logs:", font=("Arial", 12, "bold"))
        self.log_label.pack(pady=5, padx=10, anchor="w")
        
        self.log_text = ctk.CTkTextbox(self.log_frame, height=300, font=("Arial", 11))
        self.log_text.pack(pady=5, padx=10, fill="both", expand=True)
        
        # Frame para botões de ação sobre logs
        self.log_button_frame = ctk.CTkFrame(self.main_frame, corner_radius=5)
        self.log_button_frame.pack(pady=5, padx=10, fill="x")
        
        self.clear_button = ctk.CTkButton(self.log_button_frame, text="🧹 Limpar Logs", command=self.limpar_logs, 
                                         font=("Arial", 12), width=120)
        self.clear_button.pack(pady=5, padx=5, side="left")
        
        self.export_button = ctk.CTkButton(self.log_button_frame, text="💾 Exportar Logs", command=self.exportar_logs, 
                                          font=("Arial", 12), width=120)
        self.export_button.pack(pady=5, padx=5, side="left")
        
        self.open_report_button = ctk.CTkButton(self.log_button_frame, text="📋 Último Relatório", command=self.abrir_ultimo_relatorio, 
                                              font=("Arial", 12), width=120)
        self.open_report_button.pack(pady=5, padx=5, side="left")
        
        self.monitorando = True
        threading.Thread(target=self.monitorar, daemon=True).start()

    def log(self, mensagem):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", mensagem + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def adicionar_pdf(self):
        try:
            arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
            if arquivo:
                # Limpa dados anteriores
                self.ultimo_produtos = []
                self.resultados_comparacao = []
                self.ultimo_relatorio = None
                
                # Limpa os logs
                self.limpar_logs()
                
                # Copia o arquivo para a pasta de pedidos
                destino = os.path.join(self.pasta, os.path.basename(arquivo))
                shutil.copy2(arquivo, destino)  # Usando copy2 em vez de move para não remover o original
                
                self.status_label.configure(text="📄 PDF adicionado!")
                self.log(f"📄 PDF adicionado: {os.path.basename(arquivo)}")
                
                # Processa o arquivo imediatamente
                self.processar_arquivo(destino, os.path.basename(arquivo))
        except Exception as e:
            self.log(f"❌ Erro ao adicionar PDF: {e}")
            self.log(f"Detalhes: {traceback.format_exc()}")

    def limpar_logs(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.status_label.configure(text="⏳ Aguardando PDF...")

    def exportar_logs(self):
        try:
            conteudo = self.log_text.get("1.0", "end")
            arquivo = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Arquivo de Texto", "*.txt")])
            if arquivo:
                with open(arquivo, "w", encoding="utf-8") as f:
                    f.write(conteudo)
                self.log(f"✅ Logs exportados para {arquivo}")
        except Exception as e:
            self.log(f"❌ Erro ao exportar logs: {e}")

    def ver_produtos(self):
        if not self.ultimo_produtos:
            self.log("⚠️ Nenhum produto processado recentemente.")
            return
        
        # Cria uma janela para mostrar os produtos
        popup = ctk.CTkToplevel(self.root)
        popup.title("Produtos Processados")
        popup.geometry("500x400")
        popup.grab_set()  # Faz a janela modal
        
        # Título
        titulo = ctk.CTkLabel(popup, text="Produtos Extraídos", font=("Arial", 16, "bold"))
        titulo.pack(pady=10)
        
        # Lista de produtos
        produtos_frame = ctk.CTkScrollableFrame(popup, width=450, height=300)
        produtos_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Adiciona cada produto à lista
        for i, produto in enumerate(self.ultimo_produtos):
            produto_text = f"{i+1}. {produto['product_name']} - Qtd: {produto['estoque_quant']}"
            produto_label = ctk.CTkLabel(produtos_frame, text=produto_text, anchor="w", justify="left")
            produto_label.pack(pady=2, padx=5, fill="x")
        
        # Botão para fechar
        fechar_btn = ctk.CTkButton(popup, text="Fechar", command=popup.destroy)
        fechar_btn.pack(pady=10)

    def abrir_pasta(self):
        try:
            os.startfile(self.pasta)
        except Exception as e:
            self.log(f"❌ Erro ao abrir pasta: {e}")

    def processar_arquivo(self, caminho_pdf, arquivo):
        try:
            self.status_label.configure(text="🔄 Processando...")
            self.log(f"📑 Processando PDF: {caminho_pdf}")
            
            texto_extraido = extrair_texto_pdf(caminho_pdf)
            if "Erro" in texto_extraido:
                self.log(texto_extraido)
                return
                
            # Salva o texto extraído em um arquivo temporário para debug
            temp_file = os.path.join(self.pasta_relatorios, "ultimo_texto_extraido.txt")
            with open(temp_file, "w", encoding="utf-8") as f:
                f.write(texto_extraido)
            self.log("📝 Texto extraído salvo para análise")
            
            produtos_processados = processar_dados(texto_extraido)
            self.ultimo_produtos = produtos_processados  # Salva os produtos para visualização posterior
            
            if produtos_processados:
                self.log(f"📦 Produtos extraídos: {len(produtos_processados)} itens")
                for prod in produtos_processados:
                    self.log(f"- {prod['product_name']}: {prod['estoque_quant']}")
                
                # Apenas comparar com o banco de dados sem atualizar
                self.comparar_com_banco(produtos_processados)
            else:
                self.log("⚠️ Nenhum produto encontrado no PDF. Verifique o formato.")
            
            self.processados.add(arquivo)
            self.status_label.configure(text="✅ Concluído!")
            self.log(f"✅ Arquivo {arquivo} processado")
            
            # Remove o arquivo PDF após processamento completo
            try:
                os.remove(caminho_pdf)
                self.log(f"🗑️ Arquivo PDF removido após processamento")
            except:
                self.log("⚠️ Não foi possível remover o arquivo PDF")
            
        except Exception as e:
            self.log(f"❌ Erro ao processar arquivo {arquivo}: {e}")
            self.log(f"Detalhes: {traceback.format_exc()}")
            self.status_label.configure(text="❌ Erro!")

    def comparar_com_banco(self, produtos_pdf):
        try:
            self.log("🔍 Comparando produtos com o banco de dados...")
            produtos_banco = obter_produtos_do_banco()
            
            if not produtos_banco:
                self.log("❌ Erro: Nenhum produto obtido do banco de dados")
                return
            
            self.log(f"📊 Total de produtos no banco: {len(produtos_banco)}")
            
            # Lista para armazenar os resultados da comparação
            resultados = []
            
            for produto_pdf in produtos_pdf:
                self.log(f"\n🔎 Buscando correspondência para: {produto_pdf['product_name']}")
                produto_banco = encontrar_produto_correspondente(produto_pdf['product_name'], produtos_banco)
                
                if produto_banco:
                    estoque_atual = produto_banco['estoque_quant'] if produto_banco['is_bulk'] == 0 else produto_banco['estoque_peso']
                    
                    # Salva o resultado da comparação
                    resultados.append({
                        'produto_pdf': produto_pdf['product_name'],
                        'produto_banco': produto_banco['product_name'],
                        'quantidade_pdf': produto_pdf['estoque_quant'],
                        'estoque_atual': estoque_atual,
                        'id_produto': produto_banco['id'],
                        'is_bulk': produto_banco['is_bulk']
                    })
                    
                    # Mostra a correspondência encontrada
                    self.log(f"✅ Correspondência encontrada: {produto_banco['product_name']}")
                    self.log(f"   Estoque atual: {estoque_atual} | Quantidade no PDF: {produto_pdf['estoque_quant']}")
                else:
                    self.log(f"❌ Nenhuma correspondência encontrada para: {produto_pdf['product_name']}")
                    # Adiciona à lista de resultados mesmo sem correspondência
                    resultados.append({
                        'produto_pdf': produto_pdf['product_name'],
                        'produto_banco': 'Não encontrado',
                        'quantidade_pdf': produto_pdf['estoque_quant'],
                        'estoque_atual': 'N/A',
                        'id_produto': 'N/A',
                        'is_bulk': 0
                    })
            
            # Armazena os resultados para uso posterior
            self.resultados_comparacao = resultados
            
            # Salva os resultados da comparação
            self.ultimo_relatorio = self.salvar_resultados_comparacao(resultados)
            
            # Contabiliza os resultados
            encontrados = sum(1 for r in resultados if r['produto_banco'] != 'Não encontrado')
            self.log(f"\n📋 Resumo da comparação:")
            self.log(f"   Total de produtos no PDF: {len(produtos_pdf)}")
            self.log(f"   Produtos encontrados no banco: {encontrados}")
            self.log(f"   Produtos não encontrados: {len(produtos_pdf) - encontrados}")
            
        except Exception as e:
            self.log(f"❌ Erro ao comparar com o banco de dados: {e}")

    def salvar_resultados_comparacao(self, resultados):
        try:
            # Cria a pasta para relatórios se não existir
            if not os.path.exists(self.pasta_relatorios):
                os.makedirs(self.pasta_relatorios)
            
            # Nome do arquivo com data e hora
            data_hora = time.strftime("%Y%m%d_%H%M%S")
            nome_arquivo = os.path.join(self.pasta_relatorios, f"comparacao_{data_hora}.txt")
            
            with open(nome_arquivo, "w", encoding="utf-8") as f:
                f.write("RESULTADOS DA COMPARAÇÃO DE PRODUTOS\n")
                f.write("==================================\n\n")
                f.write(f"Data e hora: {time.strftime('%d/%m/%Y %H:%M:%S')}\n\n")
                
                for i, resultado in enumerate(resultados, 1):
                    f.write(f"{i}. Produto PDF: {resultado['produto_pdf']}\n")
                    f.write(f"   Correspondência: {resultado['produto_banco']}\n")
                    f.write(f"   Quantidade no PDF: {resultado['quantidade_pdf']}\n")
                    f.write(f"   Estoque atual: {resultado['estoque_atual']}\n")
                    f.write(f"   ID: {resultado['id_produto']}\n\n")
                
                # Resumo
                encontrados = sum(1 for r in resultados if r['produto_banco'] != 'Não encontrado')
                f.write(f"RESUMO:\n")
                f.write(f"Total de produtos analisados: {len(resultados)}\n")
                f.write(f"Produtos encontrados no banco: {encontrados}\n")
                f.write(f"Produtos não encontrados: {len(resultados) - encontrados}\n")
            
            self.log(f"✅ Resultados da comparação salvos em: {os.path.basename(nome_arquivo)}")
            
            # Também cria uma cópia com nome "ultimo_relatorio.txt" para fácil acesso
            ultima_copia = os.path.join(self.pasta_relatorios, f"ultimo_relatorio.txt")
            shutil.copy2(nome_arquivo, ultima_copia)
            
            return nome_arquivo
        except Exception as e:
            self.log(f"❌ Erro ao salvar resultados da comparação: {e}")
            self.log(f"Detalhes: {traceback.format_exc()}")
            return None

    def ver_relatorio(self):
        if not self.resultados_comparacao:
            self.log("⚠️ Nenhuma comparação realizada recentemente.")
            return
        
        # Cria uma janela para mostrar o relatório
        popup = ctk.CTkToplevel(self.root)
        popup.title("Relatório de Comparação")
        popup.geometry("700x500")
        popup.grab_set()  # Faz a janela modal
        
        # Título
        titulo = ctk.CTkLabel(popup, text="Relatório de Comparação com Banco de Dados", font=("Arial", 16, "bold"))
        titulo.pack(pady=10)
        
        # Frame para a tabela
        tabela_frame = ctk.CTkScrollableFrame(popup, width=650, height=350)
        tabela_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Cabeçalho
        header_frame = ctk.CTkFrame(tabela_frame)
        header_frame.pack(fill="x", pady=5)
        
        # Colunas do cabeçalho
        colunas = ["Produto PDF", "Correspondência", "Qtd PDF", "Estoque Atual"]
        larguras = [250, 250, 70, 100]
        
        for i, coluna in enumerate(colunas):
            label = ctk.CTkLabel(header_frame, text=coluna, font=("Arial", 12, "bold"), width=larguras[i])
            label.grid(row=0, column=i, padx=5)
        
        # Adiciona cada item à tabela
        for i, resultado in enumerate(self.resultados_comparacao):
            row_frame = ctk.CTkFrame(tabela_frame)
            row_frame.pack(fill="x", pady=2)
            
            # Define a cor de fundo com base na correspondência
            bg_color = "#2f2f2f" if i % 2 == 0 else "#3f3f3f"
            
            # Produto PDF
            nome_pdf = ctk.CTkLabel(row_frame, text=resultado['produto_pdf'], anchor="w", width=larguras[0])
            nome_pdf.grid(row=0, column=0, padx=5)
            
            # Produto Banco
            nome_banco = ctk.CTkLabel(row_frame, text=resultado['produto_banco'], anchor="w", width=larguras[1])
            nome_banco.grid(row=0, column=1, padx=5)
            
            # Quantidade PDF
            qtd_pdf = ctk.CTkLabel(row_frame, text=str(resultado['quantidade_pdf']), width=larguras[2])
            qtd_pdf.grid(row=0, column=2, padx=5)
            
            # Estoque Atual
            estoque = ctk.CTkLabel(row_frame, text=str(resultado['estoque_atual']), width=larguras[3])
            estoque.grid(row=0, column=3, padx=5)
        
        # Resumo
        encontrados = sum(1 for r in self.resultados_comparacao if r['produto_banco'] != 'Não encontrado')
        
        resumo_frame = ctk.CTkFrame(popup)
        resumo_frame.pack(fill="x", pady=10, padx=10)
        
        resumo_texto = (
            f"Total de produtos analisados: {len(self.resultados_comparacao)}\n"
            f"Produtos encontrados no banco: {encontrados}\n"
            f"Produtos não encontrados: {len(self.resultados_comparacao) - encontrados}"
        )
        
        resumo_label = ctk.CTkLabel(resumo_frame, text=resumo_texto, font=("Arial", 12))
        resumo_label.pack(pady=5)
        
        # Botões
        botoes_frame = ctk.CTkFrame(popup)
        botoes_frame.pack(fill="x", pady=10)
        
        fechar_btn = ctk.CTkButton(botoes_frame, text="Fechar", command=popup.destroy, width=120)
        fechar_btn.grid(row=0, column=0, padx=10, pady=5)
        
        atualizar_btn = ctk.CTkButton(botoes_frame, text="Atualizar Estoque", command=lambda: [self.atualizar_estoque_manual(), popup.destroy()], width=150)
        atualizar_btn.grid(row=0, column=1, padx=10, pady=5)
        
        abrir_relatorio_btn = ctk.CTkButton(botoes_frame, text="Abrir Relatório", command=self.abrir_ultimo_relatorio, width=120)
        abrir_relatorio_btn.grid(row=0, column=2, padx=10, pady=5)
    
    def abrir_ultimo_relatorio(self):
        if self.ultimo_relatorio and os.path.exists(self.ultimo_relatorio):
            try:
                os.startfile(self.ultimo_relatorio)
            except Exception as e:
                self.log(f"❌ Erro ao abrir relatório: {e}")
        else:
            self.log("⚠️ Nenhum relatório disponível para abrir.")
    
    def atualizar_estoque_manual(self):
        try:
            ultimo_relatorio = self.obter_ultimo_relatorio()
            if not ultimo_relatorio:
                self.log("❌ Nenhum relatório de comparação encontrado.")
                messagebox.showerror("Erro", "Nenhum relatório de comparação encontrado.")
                return

            self.log(f"📂 Lendo relatório: {os.path.basename(ultimo_relatorio)}")
            produtos_para_atualizar = []
            
            try:
                with open(ultimo_relatorio, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                    
                    # Extrai cada entrada numerada do relatório
                    entradas = re.findall(r'\d+\.\s+Produto PDF:\s+(.*?)\n\s+Correspondência:\s+(.*?)\n\s+Quantidade no PDF:\s+(\d+(?:\.\d+)?)\n\s+Estoque atual:', conteudo, re.DOTALL)
                    
                    for produto_pdf, correspondencia, qtd in entradas:
                        # Verifica se o produto tem correspondência válida
                        if correspondencia and correspondencia.strip() != "Não encontrado":
                            try:
                                quantidade = int(float(qtd))
                                produtos_para_atualizar.append({
                                    'produto_pdf': produto_pdf.strip(),
                                    'correspondencia': correspondencia.strip(), 
                                    'quantidade': quantidade
                                })
                                self.log(f"✅ Produto para atualizar: {produto_pdf.strip()} → {correspondencia.strip()} (Qtd: {quantidade})")
                            except ValueError:
                                self.log(f"⚠️ Quantidade inválida para {produto_pdf}: {qtd}")
            except Exception as e:
                self.log(f"❌ Erro ao ler arquivo de relatório: {e}")
                self.log(f"Detalhes: {traceback.format_exc()}")
                messagebox.showerror("Erro", f"Erro ao ler arquivo de relatório: {e}")
                return
            
            # Filtra apenas produtos com correspondência
            produtos_validos = [p for p in produtos_para_atualizar if p["correspondencia"] and p["quantidade"]]
            
            if not produtos_validos:
                self.log("❌ Nenhum produto válido para atualização encontrado no relatório.")
                messagebox.showerror("Erro", "Nenhum produto válido para atualização encontrado no relatório.")
                return
            
            # Confirma com o usuário
            resultado = messagebox.askyesno(
                "Confirmar Atualização", 
                f"Deseja atualizar o estoque com {len(produtos_validos)} produtos?\n\n" +
                "\n".join([f"• {p['produto_pdf']} → {p['correspondencia']} (Qtd: {p['quantidade']})" for p in produtos_validos[:5]]) +
                ("\n• ..." if len(produtos_validos) > 5 else "")
            )
            
            if not resultado:
                self.log("⏹️ Atualização de estoque cancelada pelo usuário.")
                return
                
            # Atualiza o estoque
            try:
                produtos_banco = obter_produtos_do_banco()
                conexao = mysql.connector.connect(
                    host="autorack.proxy.rlwy.net",
                    user="root",
                    password="AGWseadASVhFzAaAlxmLBoYBzgvBQhVT",
                    database="railway",
                    port=16717
                )
                
                if conexao.is_connected():
                    cursor = conexao.cursor()
                    self.log("🔄 Iniciando atualização do estoque...")
                    produtos_atualizados = 0
                    
                    for produto in produtos_validos:
                        # Encontra o produto correspondente no banco
                        produto_banco = None
                        for p in produtos_banco:
                            if p['product_name'].lower() == produto['correspondencia'].lower():
                                produto_banco = p
                                break
                        
                        if not produto_banco:
                            self.log(f"⚠️ Produto não encontrado no banco: {produto['correspondencia']}")
                            continue
                        
                        # Atualiza o estoque
                        if produto_banco['is_bulk'] == 1:
                            query = "UPDATE produto SET estoque_peso = estoque_peso + %s WHERE id = %s"
                            valor_adicionar = float(produto['quantidade'])
                            cursor.execute(query, (valor_adicionar, produto_banco['id']))
                            novo_estoque = produto_banco['estoque_peso'] + valor_adicionar
                            tipo_estoque = "peso"
                        else:
                            query = "UPDATE produto SET estoque_quant = estoque_quant + %s WHERE id = %s"
                            valor_adicionar = int(produto['quantidade'])
                            cursor.execute(query, (valor_adicionar, produto_banco['id']))
                            novo_estoque = produto_banco['estoque_quant'] + valor_adicionar
                            tipo_estoque = "quantidade"
                        
                        self.log(f"✅ Atualizado: {produto_banco['product_name']} | Estoque {tipo_estoque} +{valor_adicionar} = {novo_estoque}")
                        produtos_atualizados += 1
                    
                    conexao.commit()
                    self.log(f"✅ {produtos_atualizados} produtos atualizados com sucesso!")
                    messagebox.showinfo("Sucesso", f"{produtos_atualizados} produtos atualizados com sucesso!")
                    
                    # Limpa dados após atualização bem-sucedida
                    self.ultimo_produtos = []
                    self.resultados_comparacao = []
                    
            except Error as e:
                self.log(f"❌ Erro ao conectar ao MySQL: {e}")
                self.log(f"Detalhes: {traceback.format_exc()}")
                messagebox.showerror("Erro", f"Erro de conexão com o banco de dados: {e}")
            finally:
                if 'conexao' in locals() and conexao.is_connected():
                    cursor.close()
                    conexao.close()
                    self.log("Conexão ao MySQL fechada.")
        except Exception as e:
            self.log(f"❌ Erro ao atualizar estoque: {e}")
            self.log(f"Detalhes: {traceback.format_exc()}")
            messagebox.showerror("Erro", f"Erro ao atualizar estoque: {e}")

    def monitorar(self):
        while self.monitorando:
            try:
                arquivos_encontrados = os.listdir(self.pasta)
                for arquivo in arquivos_encontrados:
                    if arquivo.endswith(".pdf") and arquivo not in self.processados:
                        caminho_pdf = os.path.join(self.pasta, arquivo)
                        self.processar_arquivo(caminho_pdf, arquivo)
                self.root.after(5000, self.monitorar)  # Chama novamente após 5 segundos
                break  # Sai do loop, o after cuida da repetição
            except Exception as e:
                self.log(f"❌ Erro no monitoramento: {e}")
                self.status_label.configure(text="❌ Erro!")
                time.sleep(5)

    def on_closing(self):
        self.monitorando = False
        self.root.destroy()

    def obter_ultimo_relatorio(self):
        pasta_relatorios = os.path.join(os.path.dirname(os.path.abspath(__file__)), "relatorios")
        if not os.path.exists(pasta_relatorios):
            return None
        
        arquivos = os.listdir(pasta_relatorios)
        arquivos_txt = [os.path.join(pasta_relatorios, f) for f in arquivos if f.endswith(".txt")]
        if not arquivos_txt:
            return None
        
        # Retorna o arquivo mais recente
        return max(arquivos_txt, key=os.path.getmtime)

    def limpar_dados(self):
        """Limpa todos os dados temporários do processamento"""
        self.ultimo_produtos = []
        self.resultados_comparacao = []
        self.ultimo_relatorio = None
        self.log("🧹 Dados temporários limpos")

if __name__ == "__main__":
    root = ctk.CTk()
    app = PDFBotApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop() 