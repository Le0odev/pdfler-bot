import fitz  # PyMuPDF
import re
import os
import mysql.connector
from mysql.connector import Error
import time
from difflib import SequenceMatcher
import customtkinter as ctk
from tkinter import filedialog, messagebox
import shutil
import threading
import pandas as pd
import tabula
import io
import tempfile

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def extrair_texto_pdf(caminho_pdf):
    """Extrai o texto completo de um arquivo PDF."""
    try:
        texto_completo = ""
        with fitz.open(caminho_pdf) as pdf:
            for pagina in pdf:
                texto_completo += pagina.get_text("text") + "\n"
        return texto_completo
    except Exception as e:
        return f"Erro ao extrair texto do PDF: {e}"

def extrair_tabela_pdf(caminho_pdf):
    """Tenta extrair tabelas do PDF usando tabula-py."""
    try:
        # Tenta extrair todas as tabelas do PDF
        tabelas = tabula.read_pdf(caminho_pdf, pages='all', multiple_tables=True)
        return tabelas
    except Exception as e:
        print(f"Erro ao extrair tabelas com tabula: {e}")
        return []

def processar_dados_padrao(texto):
    """Processa os dados no formato padrão (como no bot_rg.py)."""
    pedidos = []
    if isinstance(texto, str) and "Erro" in texto:
        return pedidos
    
    linhas = texto.split("\n")
    item_atual = []
    
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue
        
        if re.match(r"^\d+$", linha):
            if len(item_atual) >= 8:
                try:
                    nome = item_atual[1].strip()
                    quantidade = float(item_atual[3].replace(",", "."))
                    valor = float(item_atual[4].replace(",", "."))
                    unidade = item_atual[6].strip()
                    is_bulk = 1 if unidade == "KG" else 0
                    
                    if is_bulk:
                        estoque_peso = quantidade
                        estoque_quant = 0
                    else:
                        estoque_peso = 0.00
                        estoque_quant = int(quantidade)
                    
                    pedidos.append({
                        "product_name": nome,
                        "estoque_quant": estoque_quant,
                        "estoque_peso": estoque_peso,
                        "valor": valor,
                        "is_bulk": is_bulk
                    })
                except (ValueError, TypeError) as e:
                    print(f"Erro ao processar item {item_atual}: {e}")
            
            item_atual = [linha]
        else:
            item_atual.append(linha)
    
    # Processar o último item
    if len(item_atual) >= 8:
        try:
            nome = item_atual[1].strip()
            quantidade = float(item_atual[3].replace(",", "."))
            valor = float(item_atual[4].replace(",", "."))
            unidade = item_atual[6].strip()
            is_bulk = 1 if unidade == "KG" else 0
            
            if is_bulk:
                estoque_peso = quantidade
                estoque_quant = 0
            else:
                estoque_peso = 0.00
                estoque_quant = int(quantidade)
            
            pedidos.append({
                "product_name": nome,
                "estoque_quant": estoque_quant,
                "estoque_peso": estoque_peso,
                "valor": valor,
                "is_bulk": is_bulk
            })
        except (ValueError, TypeError) as e:
            print(f"Erro ao processar item {item_atual}: {e}")
    
    return pedidos

def extrair_padroes_produto(texto):
    """Extrai produtos usando padrões de regex para diferentes formatos."""
    produtos = []
    
    # Padrão para formato "Item Código Descrição Quantidade Preço"
    padrao_tabular = r'(\d+)\s+(\w+)\s+(.*?)\s+(\d+\s*(?:UN|KG|g|ml|L))\s+R\$\s*(\d+[.,]\d+)'
    matches = re.finditer(padrao_tabular, texto, re.MULTILINE)
    
    for match in matches:
        try:
            nome_produto = match.group(3).strip()
            quantidade_str = match.group(4).strip()
            valor_str = match.group(5).strip()
            
            # Extrai quantidade e unidade
            qtd_match = re.search(r'(\d+)\s*(UN|KG|g|ml|L)', quantidade_str, re.IGNORECASE)
            if qtd_match:
                quantidade = int(qtd_match.group(1))
                unidade = qtd_match.group(2).upper()
            else:
                quantidade = int(re.search(r'(\d+)', quantidade_str).group(1))
                unidade = "UN"
            
            # Determina se é granel
            is_bulk = 1 if unidade in ["KG", "G", "ML", "L"] else 0
            
            # Calcula estoque
            if is_bulk:
                estoque_peso = float(quantidade)
                estoque_quant = 0
            else:
                estoque_peso = 0.00
                estoque_quant = quantidade
            
            # Valor
            valor = float(valor_str.replace(',', '.'))
            
            produtos.append({
                "product_name": nome_produto,
                "estoque_quant": estoque_quant,
                "estoque_peso": estoque_peso,
                "valor": valor,
                "is_bulk": is_bulk
            })
        except Exception as e:
            print(f"Erro ao processar item com padrão tabular: {e}")
    
    # Padrão para produtos em linhas separadas (Nome: quantidade, valor)
    if not produtos:
        linhas = texto.split('\n')
        produto_atual = None
        quantidade_atual = None
        
        for linha in linhas:
            # Procura por nome de produto no início da linha
            nome_match = re.match(r'^(.{10,50})\s*$', linha.strip())
            if nome_match and not re.search(r'CNPJ|CEP|Telefone|Total|Subtotal', linha, re.IGNORECASE):
                produto_atual = nome_match.group(1).strip()
                continue
                
            # Procura por quantidade e preço
            if produto_atual:
                # Busca padrões como "10 UN" ou "10 KG" ou "10 unidades"
                qtd_match = re.search(r'(\d+)\s*(UN|KG|g|ml|L|unidades?|quilos?)', linha, re.IGNORECASE)
                if qtd_match:
                    quantidade = int(qtd_match.group(1))
                    unidade = qtd_match.group(2).upper()
                    
                    # Procura preço na mesma linha
                    preco_match = re.search(r'R\$\s*(\d+[.,]\d+)', linha)
                    if preco_match:
                        valor = float(preco_match.group(1).replace(',', '.'))
                        
                        # Determina se é granel
                        is_bulk = 1 if unidade in ["KG", "G", "ML", "L", "QUILO", "QUILOS"] else 0
                        
                        # Calcula estoque
                        if is_bulk:
                            estoque_peso = float(quantidade)
                            estoque_quant = 0
                        else:
                            estoque_peso = 0.00
                            estoque_quant = quantidade
                        
                        produtos.append({
                            "product_name": produto_atual,
                            "estoque_quant": estoque_quant,
                            "estoque_peso": estoque_peso,
                            "valor": valor,
                            "is_bulk": is_bulk
                        })
                        produto_atual = None
    
    return produtos

def processar_tabela_extraida(tabelas):
    """Processa tabelas extraídas do PDF."""
    produtos = []
    
    if not tabelas:
        return produtos
    
    for tabela in tabelas:
        if tabela.empty:
            continue
        
        # Tentar identificar colunas relevantes
        colunas = tabela.columns.tolist()
        col_produto = None
        col_quantidade = None
        col_preco = None
        
        # Busca colunas por nome
        for col in colunas:
            col_lower = str(col).lower()
            if any(termo in col_lower for termo in ['produto', 'descrição', 'item', 'mercadoria']):
                col_produto = col
            elif any(termo in col_lower for termo in ['qtde', 'quantidade', 'qtd', 'quant']):
                col_quantidade = col
            elif any(termo in col_lower for termo in ['preço', 'valor', 'r$', 'unit']):
                col_preco = col
        
        # Se não encontrou por nome, tenta por posição
        if col_produto is None and len(colunas) >= 3:
            col_produto = colunas[2]
        if col_quantidade is None and len(colunas) >= 4:
            col_quantidade = colunas[3]
        if col_preco is None and len(colunas) >= 5:
            col_preco = colunas[4]
        
        # Processa cada linha
        for _, linha in tabela.iterrows():
            try:
                # Busca nome do produto
                nome_produto = None
                if col_produto and pd.notna(linha[col_produto]):
                    nome_produto = str(linha[col_produto]).strip()
                
                # Se não achou nome, busca em todas as colunas
                if not nome_produto:
                    for col in colunas:
                        valor = str(linha[col])
                        if len(valor) > 10 and not valor.startswith('R$') and not re.match(r'^\d+(\.\d+)?$', valor):
                            nome_produto = valor.strip()
                            break
                
                if not nome_produto:
                    continue
                
                # Busca quantidade
                quantidade = None
                unidade = "UN"
                if col_quantidade and pd.notna(linha[col_quantidade]):
                    qtd_str = str(linha[col_quantidade])
                    qtd_match = re.search(r'(\d+)\s*(UN|KG|g|ml|L)?', qtd_str, re.IGNORECASE)
                    if qtd_match:
                        quantidade = int(qtd_match.group(1))
                        if qtd_match.group(2):
                            unidade = qtd_match.group(2).upper()
                
                # Se não achou, assume 1
                if not quantidade:
                    quantidade = 1
                
                # Busca preço
                preco = None
                if col_preco and pd.notna(linha[col_preco]):
                    preco_str = str(linha[col_preco])
                    preco_match = re.search(r'R\$\s*(\d+[.,]\d+)', preco_str)
                    if preco_match:
                        preco = float(preco_match.group(1).replace(',', '.'))
                    else:
                        # Tenta converter diretamente
                        try:
                            preco_str = preco_str.replace('R$', '').replace('.', '').replace(',', '.')
                            preco = float(preco_str)
                        except:
                            pass
                
                # Se não achou preço, procura em todas as colunas
                if not preco:
                    for col in colunas:
                        valor = str(linha[col])
                        preco_match = re.search(r'R\$\s*(\d+[.,]\d+)', valor)
                        if preco_match:
                            preco = float(preco_match.group(1).replace(',', '.'))
                            break
                
                # Se ainda não achou, assume 0
                if not preco:
                    preco = 0.0
                
                # Determina se é granel
                is_bulk = 1 if unidade in ["KG", "G", "ML", "L"] else 0
                
                # Calcula estoque
                if is_bulk:
                    estoque_peso = float(quantidade)
                    estoque_quant = 0
                else:
                    estoque_peso = 0.00
                    estoque_quant = quantidade
                
                produtos.append({
                    "product_name": nome_produto,
                    "estoque_quant": estoque_quant,
                    "estoque_peso": estoque_peso,
                    "valor": preco,
                    "is_bulk": is_bulk
                })
            except Exception as e:
                print(f"Erro ao processar linha da tabela: {e}")
    
    return produtos

def processar_dados_completos(caminho_pdf):
    """Processa o PDF usando múltiplos métodos para maior compatibilidade."""
    # Extrai texto e tabelas
    texto = extrair_texto_pdf(caminho_pdf)
    tabelas = extrair_tabela_pdf(caminho_pdf)
    
    # Lista para armazenar todos os produtos encontrados
    todos_produtos = []
    
    # Tenta o método padrão
    if isinstance(texto, str) and "Erro" not in texto:
        produtos_padrao = processar_dados_padrao(texto)
        if produtos_padrao:
            todos_produtos.extend(produtos_padrao)
    
    # Tenta extrair com padrões de regex
    if isinstance(texto, str) and "Erro" not in texto and not todos_produtos:
        produtos_regex = extrair_padroes_produto(texto)
        if produtos_regex:
            todos_produtos.extend(produtos_regex)
    
    # Tenta processar tabelas
    if not todos_produtos and tabelas:
        produtos_tabela = processar_tabela_extraida(tabelas)
        if produtos_tabela:
            todos_produtos.extend(produtos_tabela)
    
    # Retorna os produtos encontrados e o texto para debug
    return todos_produtos, texto

def obter_produtos_do_banco():
    """Obtém a lista de produtos do banco de dados."""
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
    """Normaliza o texto para comparação, removendo acentos, caracteres especiais, etc."""
    import unicodedata
    import re
    texto = texto.lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = re.sub(r'[^\w\s]', ' ', texto)
    
    # Lista de marcas comuns de suplementos para padronização
    marcas = {
        'body action': 'bodyaction',
        'bodyaction': 'bodyaction',
        'body nutritions': 'bodynutritions',
        'black skull': 'blackskull',
        'blackskull': 'blackskull',
        'integral medica': 'integralmedica',
        'integralmedica': 'integralmedica',
        'integral': 'integralmedica',
        'max titanium': 'maxtitanium',
        'maxtitanium': 'maxtitanium',
        'red lion': 'redlion',
        'redlion': 'redlion',
        'darkness': 'darkness',
        'dark': 'darkness',
        'naturovos': 'naturovos',
        'demons lab': 'demonslab',
        'demonslab': 'demonslab'
    }
    
    # Procura por marcas conhecidas no texto
    texto_lower = texto.lower()
    for marca, padrao in marcas.items():
        if marca in texto_lower:
            # Substitui a marca pela versão padronizada
            texto_lower = texto_lower.replace(marca, padrao)
    
    # Continua com a normalização normal
    substituicoes = {'org': 'organico', 'ml': '', 'g': '', 'kg': '', 'grs': '', 'gr': '', 'lt': '', 'l': '',
                     'pct': '', 'un': '', 'cx': '', 'c/': 'com', 's/': 'sem', 
                     ' de ': ' ', ' da ': ' ', ' do ': ' ', ' com ': ' ', ' e ': ' '}
    for abrev, completo in substituicoes.items():
        texto_lower = re.sub(r'\b' + abrev + r'\b' if len(abrev) <= 4 else abrev, completo, texto_lower)
    
    palavras_ignorar = ['de', 'da', 'do', 'em', 'a', 'o', 'e', 'com', 'sem', 'para', 'em', 'no', 'na', 'os', 'as']
    palavras = texto_lower.split()
    palavras_filtradas = [p for p in palavras if p not in palavras_ignorar]
    palavras_sem_numeros = [p for p in palavras_filtradas if not re.match(r'^\d+$', p)]
    
    # Trata termos específicos de suplementos
    suplementos_termos = {
        'whey': 'wheyprotein',
        'protein': 'wheyprotein',
        'wheyprotein': 'wheyprotein',
        'creatina': 'creatina',
        'creatine': 'creatina',
        'glutamina': 'glutamina',
        'glutamine': 'glutamina',
        'albumina': 'albumina',
        'albumin': 'albumina',
        'pre': 'pretraining',
        'preworkout': 'pretraining',
        'pretraining': 'pretraining',
        'pre workout': 'pretraining',
        'pre treino': 'pretraining',
        'pretreino': 'pretraining',
        'termogenico': 'termogenico',
        'thermo': 'termogenico',
        'abdomem': 'abdomen',
        'isolado': 'isolate',
        'isolate': 'isolate',
        'concentrated': 'concentrate',
        'concentrate': 'concentrate',
        'concentrado': 'concentrate'
    }
    
    resultado = []
    for palavra in palavras_sem_numeros:
        adicionada = False
        for termo, padrao in suplementos_termos.items():
            if termo in palavra:
                resultado.append(padrao)
                adicionada = True
                break
        if not adicionada:
            resultado.append(palavra)
    
    return ' '.join(resultado)

def preprocessar_nome_produto(nome):
    """Pré-processa o nome do produto para melhorar as chances de correspondência."""
    # Remove gramagens e informações entre parênteses
    nome_limpo = re.sub(r'\(.*?\)', '', nome)
    nome_limpo = re.sub(r'\d+\s*[gkg]+\b', '', nome_limpo)
    nome_limpo = re.sub(r'\d+\s*[ml]+\b', '', nome_limpo)
    nome_limpo = re.sub(r'\d+\s*caps\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\d+\s*capsulas\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\d+\s*comprimidos\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\d+\s*tabletes\b', '', nome_limpo, flags=re.IGNORECASE)
    
    # Remove outros termos não relevantes
    nome_limpo = re.sub(r'sabor\s+\w+', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'embalagem', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'pacote', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'unidade', '', nome_limpo, flags=re.IGNORECASE)
    
    # Remove indicações de cores, que geralmente se referem a sabores
    nome_limpo = re.sub(r'\s+(amarelo|amarela|vermelho|vermelha|azul|roxo|roxa|verde)\b', '', nome_limpo, flags=re.IGNORECASE)
    
    # Remove indicações de sabores comuns
    nome_limpo = re.sub(r'\s+(chocolate|morango|baunilha|natural|frutas|laranja|limao|limão|maracuja|maracujá|tutti\s*frutti|uva|abacaxi)\b', '', nome_limpo, flags=re.IGNORECASE)
    
    # Remove cores dos suplementos
    nome_limpo = re.sub(r'\s+black\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\s+blue\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\s+red\b', '', nome_limpo, flags=re.IGNORECASE)
    
    # Remove tokens especiais de sabores pra produtos específicos
    nome_limpo = re.sub(r'\s+f\.?\s+(amarela|roxa|vermelha)\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\s+fruit\s+punch\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\s+original\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\s+hardcore\b', '', nome_limpo, flags=re.IGNORECASE)
    nome_limpo = re.sub(r'\s+fruint\s+punch\b', '', nome_limpo, flags=re.IGNORECASE)
    
    # Normaliza produtos específicos mencionados pelo cliente
    nome_lower = nome_limpo.lower()
    
    # Albumina (todos os sabores devem cair na mesma categoria)
    if 'albumina' in nome_lower and ('naturovos' in nome_lower or 'naturo vos' in nome_lower):
        nome_limpo = "NATUROVOS ALBUMINA"
    
    # C4 (produtos da new millen)
    if 'c4' in nome_lower and ('new millen' in nome_lower or 'newmillen' in nome_lower):
        if 'beta pump' in nome_lower:
            nome_limpo = "C4 BETA PUMP NEW MILLEN"
        elif 'caffeine' in nome_lower:
            nome_limpo = "C4 CAFFEINE NEW MILLEN"
        else:
            nome_limpo = "C4 NEW MILLEN"
    
    # Ectoplasma (demons lab)
    if ('ectoplasma' in nome_lower or 'ecto plasma' in nome_lower) and ('demons' in nome_lower or 'demon' in nome_lower):
        nome_limpo = "ECTOPLASMA DEMONS LAB"
    
    # Insane (demons lab)
    if 'insane' in nome_lower and ('demons' in nome_lower or 'demon' in nome_lower):
        nome_limpo = "INSANE ORIGINAL DEMONS LAB"
    
    # Dr Peanut pasta de amendoim
    if 'pasta' in nome_lower and 'amendoim' in nome_lower and ('dr' in nome_lower or 'dr.' in nome_lower):
        nome_limpo = "PASTA DE AMENDOIM DR PEANUT"
    
    # Remove múltiplos espaços e espaços no início/fim
    nome_limpo = re.sub(r'\s+', ' ', nome_limpo).strip()
    
    return nome_limpo

def calcular_similaridade_produtos(nome1, nome2):
    """Calcula a similaridade entre dois nomes de produtos."""
    # Normaliza os nomes
    nome1_norm = normalizar_texto(nome1)
    nome2_norm = normalizar_texto(nome2)
    
    # Similarity ratio por sequência completa
    seq_similarity = SequenceMatcher(None, nome1_norm, nome2_norm).ratio()
    
    # Quebra em palavras
    palavras1 = set(nome1_norm.split())
    palavras2 = set(nome2_norm.split())
    
    # Se um dos conjuntos estiver vazio, retorna apenas o ratio da sequência
    if not palavras1 or not palavras2:
        return SequenceMatcher(None, nome1.lower(), nome2.lower()).ratio()
    
    # Obtém as palavras comuns
    palavras_comuns = palavras1.intersection(palavras2)
    
    # Se não existirem palavras comuns, retorna apenas o ratio da sequência
    if not palavras_comuns:
        return seq_similarity
    
    # Calcula o índice Jaccard (proporção de palavras comuns em relação ao total)
    jaccard = len(palavras_comuns) / len(palavras1.union(palavras2))
    
    # Calcula o coeficiente de palavras importantes
    # Identifica palavras que provavelmente são mais importantes (marcas, substantivos)
    palavras_importantes = set()
    for palavra in palavras1.union(palavras2):
        # Palavras maiores são provavelmente mais importantes (marcas, tipos específicos)
        if len(palavra) > 4:
            palavras_importantes.add(palavra)
    
    # Se temos palavras importantes, verifica se elas são comuns
    palavras_importantes_comuns = palavras_importantes.intersection(palavras_comuns)
    coef_importantes = 1.0
    if palavras_importantes and len(palavras_importantes) > 0:
        coef_importantes = len(palavras_importantes_comuns) / len(palavras_importantes)
    
    # Dá peso maior para correspondências de palavras inteiras (Jaccard) e palavras importantes
    return (seq_similarity * 0.3) + (jaccard * 0.4) + (coef_importantes * 0.3)

def encontrar_produto_correspondente(nome_produto, produtos_banco):
    """Encontra o produto correspondente no banco de dados."""
    melhor_correspondencia = None
    maior_similaridade = 0.45  # Diminuir para 0.45 para ser menos restritivo (era 0.6)
    
    # Pré-processa o nome para melhorar a correspondência
    nome_produto_limpo = preprocessar_nome_produto(nome_produto)
    
    # Lista de produtos encontrados para processamento posterior
    matches_possiveis = []
    
    # Primeiro tenta encontrar correspondência exata (ignorando case)
    for produto in produtos_banco:
        produto_nome_limpo = preprocessar_nome_produto(produto['product_name'])
        if nome_produto_limpo.lower() == produto_nome_limpo.lower():
            # Correspondência exata, retorna imediatamente
            return produto
    
    # Algumas correspondências exatas específicas que precisam de atenção especial
    nome_lower = nome_produto_limpo.lower()
    
    # Procura produtos especiais (ECTOPLASMA, etc)
    if 'ectoplasma' in nome_lower:
        for produto in produtos_banco:
            if 'ectoplasma' in produto['product_name'].lower():
                return produto
    
    # C4 New Millen (vários tipos)
    if 'c4' in nome_lower and ('new millen' in nome_lower or 'newmillen' in nome_lower):
        tipo_c4 = None
        if 'beta pump' in nome_lower:
            tipo_c4 = 'beta pump'
        elif 'caffeine' in nome_lower:
            tipo_c4 = 'caffeine'
            
        for produto in produtos_banco:
            prod_lower = produto['product_name'].lower()
            if 'c4' in prod_lower and ('new millen' in prod_lower or 'newmillen' in prod_lower):
                if tipo_c4 is None or tipo_c4 in prod_lower:
                    return produto
                    
    # Extrai palavras-chave importantes do nome do produto
    palavras_chave = set(normalizar_texto(nome_produto_limpo).split())
    
    # Cria um dicionário para armazenar todos os scores para debug posterior
    scores = {}
    
    # Se não encontrar, usa similaridade
    for produto in produtos_banco:
        produto_nome_limpo = preprocessar_nome_produto(produto['product_name'])
        
        # Calcular a similaridade
        sim = calcular_similaridade_produtos(nome_produto_limpo, produto_nome_limpo)
        
        # Guarda o score para debug
        scores[produto['product_name']] = sim
        
        # Verificação adicional: tenta garantir que pelo menos 1 palavra importante esteja presente
        produto_palavras = set(normalizar_texto(produto_nome_limpo).split())
        
        # Aumenta o score para palavras importantes compartilhadas
        palavras_comuns = palavras_chave.intersection(produto_palavras)
        if len(palavras_comuns) > 0:
            # Dá mais peso quanto mais palavras importantes em comum
            sim += 0.05 * len(palavras_comuns)
        
        # Prioriza marcas específicas
        if any(marca in nome_lower and marca in produto_nome_limpo.lower() for marca in [
            'naturovos', 'integralmedica', 'blackskull', 'redlion', 'demons', 'bodyaction'
        ]):
            sim += 0.1
            
        # Prioriza produtos específicos
        if any(produto_especifico in nome_lower and produto_especifico in produto_nome_limpo.lower() 
               for produto_especifico in [
                   'albumina', 'creatina', 'whey', 'glutamina', 'pre treino', 'c4', 
                   'termogenico', 'pasta amendoim', 'ectoplasma', 'insane'
               ]):
            sim += 0.15
        
        # Verifica se este é o melhor até agora
        if sim > maior_similaridade:
            maior_similaridade = sim
            melhor_correspondencia = produto
            
        # Guarda na lista de possíveis correspondências
        if sim > 0.4:  # Limiar mínimo para considerar
            matches_possiveis.append((produto, sim))
    
    # Ordena os scores para debug
    scores_ordenados = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    print(f"Scores para '{nome_produto}' (limpo: '{nome_produto_limpo}'):")
    for nome, score in scores_ordenados[:5]:  # Mostra apenas os 5 melhores
        print(f"  - {nome}: {score:.2f}")
    
    return melhor_correspondencia

def atualizar_estoque(produtos_pdf, app=None):
    """Atualiza o estoque no banco de dados com os produtos do PDF."""
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
            produtos_atualizados = 0
            produtos_processados = set()
            for produto_pdf in produtos_pdf:
                produto_banco = encontrar_produto_correspondente(produto_pdf['product_name'], produtos_banco)
                if produto_banco and produto_banco['id'] not in produtos_processados:
                    estoque_atual_quant = produto_banco['estoque_quant']
                    estoque_atual_peso = produto_banco['estoque_peso']
                    if app:
                        app.log(f"Produto: {produto_banco['product_name']} | Estoque atual: {estoque_atual_quant}/{estoque_atual_peso} | A adicionar: {produto_pdf['estoque_quant']}/{produto_pdf['estoque_peso']}")
                    if produto_pdf['is_bulk'] == 1:
                        query = "UPDATE produto SET estoque_peso = estoque_peso + %s WHERE id = %s"
                        cursor.execute(query, (produto_pdf['estoque_peso'], produto_banco['id']))
                        novo_estoque = estoque_atual_peso + produto_pdf['estoque_peso']
                    else:
                        query = "UPDATE produto SET estoque_quant = estoque_quant + %s WHERE id = %s"
                        cursor.execute(query, (produto_pdf['estoque_quant'], produto_banco['id']))
                        novo_estoque = estoque_atual_quant + produto_pdf['estoque_quant']
                    produtos_atualizados += 1
                    produtos_processados.add(produto_banco['id'])
                    if app:
                        app.log(f"Atualizado: {produto_banco['product_name']} | Novo estoque: {novo_estoque}")
            conexao.commit()
            if app:
                app.log(f"✅ {produtos_atualizados} produtos atualizados com sucesso.")
            return produtos_atualizados
    except Error as e:
        if app:
            app.log(f"❌ Erro ao conectar ao MySQL: {e}")
        return 0
    finally:
        if 'conexao' in locals() and conexao.is_connected():
            cursor.close()
            conexao.close()
            if app:
                app.log("Conexão ao MySQL fechada.")

# Adiciona a classe principal do aplicativo ao final do arquivo
class PDFlerUniversalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFler Bot Universal")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        self.pasta = os.path.join(os.path.expanduser("~"), "pdfler-bot", "pedidos")
        if not os.path.exists(self.pasta):
            os.makedirs(self.pasta)
        
        self.processados_dir = os.path.join(self.pasta, "processados")
        if not os.path.exists(self.processados_dir):
            os.makedirs(self.processados_dir)
        
        self.processados = set()
        self.texto_extraido = ""
        
        self.criar_interface()
        
        # Inicia o monitoramento em segundo plano
        self.monitorando = True
        self.monitor_thread = threading.Thread(target=self.monitorar, daemon=True)
        self.monitor_thread.start()
    
    def criar_interface(self):
        """Cria a interface do aplicativo com customtkinter."""
        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=10)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Título
        self.title_label = ctk.CTkLabel(self.main_frame, text="PDFler Bot Universal", font=("Arial", 20, "bold"))
        self.title_label.pack(pady=10)
        
        # Frame para botões superiores
        self.top_button_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.top_button_frame.pack(pady=5, fill="x")
        
        # Botão para adicionar PDF
        self.add_button = ctk.CTkButton(self.top_button_frame, text="Adicionar PDF", 
                                       command=self.adicionar_pdf, font=("Arial", 14), 
                                       width=200, height=40)
        self.add_button.pack(pady=10, side="left", padx=10)
        
        # Botão para processar pasta
        self.process_folder_button = ctk.CTkButton(self.top_button_frame, text="Processar Pasta", 
                                                 command=self.processar_pasta, font=("Arial", 14), 
                                                 width=200, height=40)
        self.process_folder_button.pack(pady=10, side="right", padx=10)
        
        # Label de status
        self.status_label = ctk.CTkLabel(self.main_frame, text="⏳ Aguardando PDF...", font=("Arial", 12))
        self.status_label.pack(pady=5)
        
        # Frame para debug
        self.debug_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.debug_frame.pack(pady=5, fill="x")
        
        self.debug_label = ctk.CTkLabel(self.debug_frame, text="Opções de Debug:", font=("Arial", 12, "bold"))
        self.debug_label.pack(pady=5, side="left", padx=10)
        
        self.debug_var = ctk.BooleanVar(value=True)
        self.debug_checkbox = ctk.CTkCheckBox(self.debug_frame, text="Mostrar detalhes", variable=self.debug_var)
        self.debug_checkbox.pack(pady=5, side="left", padx=10)
        
        self.save_text_button = ctk.CTkButton(self.debug_frame, text="Salvar Texto Extraído", 
                                            command=self.salvar_texto_extraido, font=("Arial", 12), 
                                            width=150)
        self.save_text_button.pack(pady=5, side="right", padx=10)
        
        # Área de log
        self.log_text = ctk.CTkTextbox(self.main_frame, height=300, width=650, font=("Arial", 11), state="disabled")
        self.log_text.pack(pady=10, fill="both", expand=True)
        
        # Frame para botões inferiores
        self.button_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.button_frame.pack(pady=5, fill="x")
        
        self.clear_button = ctk.CTkButton(self.button_frame, text="Limpar Logs", 
                                        command=self.limpar_logs, font=("Arial", 12), 
                                        width=150)
        self.clear_button.pack(pady=5, side="left", padx=10)
        
        self.open_folder_button = ctk.CTkButton(self.button_frame, text="Abrir Pasta", 
                                              command=self.abrir_pasta, font=("Arial", 12), 
                                              width=150)
        self.open_folder_button.pack(pady=5, side="right", padx=10)
        
        # Verifica dependências
        try:
            import tabula
            self.log(f"✅ Tabula-py encontrado. Extração de tabelas habilitada.")
        except ImportError:
            messagebox.showwarning("Dependência Faltando", 
                                 "A biblioteca tabula-py não está instalada. A extração de tabelas não funcionará.\n\n"
                                 "Instale com: pip install tabula-py")
            self.log(f"⚠️ Tabula-py não encontrado. Extração de tabelas desabilitada.")
    
    def log(self, mensagem):
        """Adiciona uma mensagem ao log."""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", mensagem + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
    
    def adicionar_pdf(self):
        """Adiciona um arquivo PDF à pasta de processamento."""
        try:
            arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
            if arquivo:
                destino = os.path.join(self.pasta, os.path.basename(arquivo))
                shutil.copy2(arquivo, destino)  # Usa copy2 em vez de move para manter o original
                self.status_label.configure(text="📄 PDF adicionado!")
                self.log(f"📄 PDF adicionado: {os.path.basename(arquivo)}")
        except Exception as e:
            self.log(f"❌ Erro ao adicionar PDF: {e}")
    
    def limpar_logs(self):
        """Limpa o log de mensagens."""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.status_label.configure(text="⏳ Aguardando PDF...")
    
    def abrir_pasta(self):
        """Abre a pasta de monitoramento no explorador de arquivos."""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(self.pasta)
            elif os.name == 'posix':  # macOS e Linux
                import subprocess
                subprocess.Popen(['xdg-open', self.pasta])
        except Exception as e:
            self.log(f"❌ Erro ao abrir pasta: {e}")
    
    def salvar_texto_extraido(self):
        """Salva o texto extraído em um arquivo de texto para análise."""
        if not self.texto_extraido:
            self.log("⚠️ Nenhum texto extraído para salvar.")
            return
            
        try:
            arquivo = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Arquivos de Texto", "*.txt")],
                initialdir=self.pasta
            )
            if arquivo:
                with open(arquivo, 'w', encoding='utf-8') as f:
                    f.write(self.texto_extraido)
                self.log(f"✅ Texto extraído salvo em: {arquivo}")
        except Exception as e:
            self.log(f"❌ Erro ao salvar texto: {e}")
    
    def processar_pasta(self):
        """Processa todos os PDFs na pasta manualmente."""
        try:
            arquivos = [f for f in os.listdir(self.pasta) if f.endswith('.pdf')]
            if not arquivos:
                self.log("📂 Nenhum PDF encontrado na pasta.")
                return
                
            self.log(f"📂 Processando {len(arquivos)} arquivos...")
            for arquivo in arquivos:
                caminho_pdf = os.path.join(self.pasta, arquivo)
                self.processar_arquivo(caminho_pdf, arquivo)
        except Exception as e:
            self.log(f"❌ Erro ao processar pasta: {e}")
    
    def processar_arquivo(self, caminho_pdf, arquivo):
        """Processa um arquivo PDF individual usando os métodos universais."""
        try:
            self.status_label.configure(text="🔄 Processando...")
            self.log(f"📑 Processando PDF: {arquivo}")
            
            # Usar o processador universal
            produtos, texto = processar_dados_completos(caminho_pdf)
            self.texto_extraido = texto  # Salva para possível debug
            
            # Mostrar debug se ativado
            if self.debug_var.get() and isinstance(texto, str):
                self.log("📝 Amostra do texto extraído (primeiros 300 caracteres):")
                self.log(texto[:300].replace('\n', ' ') + "...")
            
            if produtos:
                self.log(f"📦 Produtos extraídos: {len(produtos)} itens")
                for prod in produtos:
                    if prod['is_bulk'] == 1:
                        self.log(f"- {prod['product_name']}: {prod['estoque_peso']} KG | R$ {prod['valor']}")
                    else:
                        self.log(f"- {prod['product_name']}: {prod['estoque_quant']} UN | R$ {prod['valor']}")
                
                # Atualiza o estoque
                atualizar_estoque(produtos, self)
            else:
                self.log("⚠️ Nenhum produto encontrado no PDF.")
                self.log("💡 Tente salvar o texto extraído para análise.")
            
            self.processados.add(arquivo)
            self.status_label.configure(text="✅ Concluído!")
            self.log(f"✅ Arquivo {arquivo} processado")
            
            # Move o arquivo para a pasta de processados
            destino = os.path.join(self.processados_dir, arquivo)
            if os.path.exists(caminho_pdf):
                shutil.move(caminho_pdf, destino)
                self.log(f"📁 Arquivo {arquivo} movido para a pasta de processados")
        except Exception as e:
            self.log(f"❌ Erro ao processar arquivo {arquivo}: {e}")
            self.status_label.configure(text="❌ Erro!")
    
    def monitorar(self):
        """Monitora a pasta em busca de novos PDFs."""
        while self.monitorando:
            try:
                arquivos_encontrados = [f for f in os.listdir(self.pasta) if f.endswith('.pdf')]
                for arquivo in arquivos_encontrados:
                    if arquivo not in self.processados:
                        caminho_pdf = os.path.join(self.pasta, arquivo)
                        # Usa o método principal para processar o arquivo
                        self.root.after(0, lambda p=caminho_pdf, a=arquivo: self.processar_arquivo(p, a))
                time.sleep(5)  # Verifica a cada 5 segundos
            except Exception as e:
                print(f"Erro no monitoramento: {e}")
                time.sleep(5)
    
    def on_closing(self):
        """Manipula o evento de fechamento da aplicação."""
        self.monitorando = False
        self.root.destroy()

if __name__ == "__main__":
    root = ctk.CTk()
    app = PDFlerUniversalApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

