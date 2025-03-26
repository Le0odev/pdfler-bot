import fitz  # PyMuPDF
import re
import os
import mysql.connector
from mysql.connector import Error
import time
from difflib import SequenceMatcher
import customtkinter as ctk
from tkinter import filedialog
import shutil
import threading

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
    pedidos = []
    if isinstance(texto, str) and "Erro" in texto:
        return pedidos  # Retorna vazio se houve erro na extração
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
    substituicoes = {'org': 'organico', 'ml': '', 'g': '', 'kg': '', 'grs': '', 'gr': '', 'lt': '', 'l': '',
                     'pct': '', 'un': '', 'cx': '', 'c/': 'com', 's/': 'sem'}
    for abrev, completo in substituicoes.items():
        texto = re.sub(r'\b' + abrev + r'\b', completo, texto)
    palavras_ignorar = ['de', 'da', 'do', 'em', 'a', 'o', 'e', 'com', 'sem', 'para']
    palavras = texto.split()
    palavras_filtradas = [p for p in palavras if p not in palavras_ignorar]
    palavras_sem_numeros = [p for p in palavras_filtradas if not re.match(r'^\d+$', p)]
    return ' '.join(palavras_sem_numeros)

def calcular_similaridade_produtos(nome1, nome2):
    nome1_norm = normalizar_texto(nome1)
    nome2_norm = normalizar_texto(nome2)
    palavras1 = set(nome1_norm.split())
    palavras2 = set(nome2_norm.split())
    if not palavras1 or not palavras2:
        return SequenceMatcher(None, nome1.lower(), nome2.lower()).ratio()
    palavras_comuns = palavras1.intersection(palavras2)
    if not palavras_comuns:
        return SequenceMatcher(None, nome1_norm, nome2_norm).ratio()
    jaccard = len(palavras_comuns) / len(palavras1.union(palavras2))
    seq_similarity = SequenceMatcher(None, nome1_norm, nome2_norm).ratio()
    return (jaccard * 0.7) + (seq_similarity * 0.3)

def encontrar_produto_correspondente(nome_produto, produtos_banco):
    melhor_correspondencia = None
    maior_similaridade = 0.5
    for produto in produtos_banco:
        if nome_produto.lower() == produto['product_name'].lower():
            return produto
    for produto in produtos_banco:
        sim = calcular_similaridade_produtos(nome_produto, produto['product_name'])
        if sim > maior_similaridade:
            maior_similaridade = sim
            melhor_correspondencia = produto
    return melhor_correspondencia

def atualizar_estoque(produtos_pdf, app=None):
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

class PDFBotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFler Bot")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        self.pasta = "C:/Users/pc/pdfler-bot/pedidos"
        if not os.path.exists(self.pasta):
            os.makedirs(self.pasta)
        
        self.processados = set()
        
        self.main_frame = ctk.CTkFrame(root, corner_radius=10)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        self.title_label = ctk.CTkLabel(self.main_frame, text="PDFler Bot", font=("Arial", 20, "bold"))
        self.title_label.pack(pady=10)
        
        self.add_button = ctk.CTkButton(self.main_frame, text="Adicionar PDF", command=self.adicionar_pdf, font=("Arial", 14), width=200, height=40)
        self.add_button.pack(pady=10)
        
        self.status_label = ctk.CTkLabel(self.main_frame, text="⏳ Aguardando PDF...", font=("Arial", 12))
        self.status_label.pack(pady=5)
        
        self.log_text = ctk.CTkTextbox(self.main_frame, height=150, width=450, font=("Arial", 11), state="disabled")
        self.log_text.pack(pady=10)
        
        self.clear_button = ctk.CTkButton(self.main_frame, text="Limpar Logs", command=self.limpar_logs, font=("Arial", 12), width=150)
        self.clear_button.pack(pady=5)
        
        self.monitorando = True
        self.monitorar()  # Chama diretamente em vez de thread pra evitar crashes no .exe

    def log(self, mensagem):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", mensagem + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def adicionar_pdf(self):
        try:
            arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
            if arquivo:
                destino = os.path.join(self.pasta, os.path.basename(arquivo))
                shutil.move(arquivo, destino)
                self.status_label.configure(text="📄 PDF adicionado!")
                self.log(f"📄 PDF adicionado: {os.path.basename(arquivo)}")
        except Exception as e:
            self.log(f"❌ Erro ao adicionar PDF: {e}")

    def limpar_logs(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.status_label.configure(text="⏳ Aguardando PDF...")

    def processar_arquivo(self, caminho_pdf, arquivo):
        try:
            self.status_label.configure(text="🔄 Processando...")
            self.log(f"📑 Novo PDF detectado: {caminho_pdf}")
            
            texto_extraido = extrair_texto_pdf(caminho_pdf)
            if "Erro" in texto_extraido:
                self.log(texto_extraido)
            else:
                produtos_processados = processar_dados(texto_extraido)
                if produtos_processados:
                    self.log(f"📦 Produtos extraídos: {len(produtos_processados)} itens")
                    for prod in produtos_processados:
                        self.log(f"- {prod['product_name']}: {prod['estoque_quant']}/{prod['estoque_peso']}")
                    atualizar_estoque(produtos_processados, self)
            
            self.processados.add(arquivo)
            self.status_label.configure(text="✅ Concluído!")
            self.log(f"✅ Arquivo {arquivo} processado")
            
            # Remove o PDF após processar
            os.remove(caminho_pdf)
            self.log(f"🗑️ Arquivo {arquivo} removido da pasta")
        except Exception as e:
            self.log(f"❌ Erro ao processar arquivo {arquivo}: {e}")
            self.status_label.configure(text="❌ Erro!")

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

if __name__ == "__main__":
    root = ctk.CTk()
    app = PDFBotApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()