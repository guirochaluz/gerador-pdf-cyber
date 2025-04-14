
import os
import webbrowser
import requests
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys

CLIENT_ID = "607986627444537"
CLIENT_SECRET = "MtKukOnsP5JvcJjGmTLzctyHCP8pUQZs"
REDIRECT_URI = "https://econtazoom.com.br"

PASTA_SAIDA = ""
EXCEL_PATH = ""
log_resultados = []

class AnuncioPDF(FPDF):
    def __init__(self, titulo, descricao, imagens):
        super().__init__()
        self.titulo = titulo
        self.descricao = descricao
        self.imagens = imagens
        self.set_auto_page_break(auto=True, margin=15)

    def header(self):
        self.set_font("Arial", 'B', 16)
        self.set_text_color(0, 70, 140)
        self.cell(0, 10, self.titulo, ln=True, align='C')
        self.ln(5)
        self.set_draw_color(0, 70, 140)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(10)

    def add_descricao(self):
        self.set_font("Arial", 'B', 12)
        self.set_text_color(0)
        self.cell(0, 10, "Descrição do Anúncio:", ln=True)
        self.set_font("Arial", size=10)
        self.multi_cell(0, 8, self.descricao)
        self.ln(10)

    def add_imagens(self):
        for i, url in enumerate(self.imagens):
            try:
                response = requests.get(url)
                image = Image.open(BytesIO(response.content)).convert("RGB")
                temp_path = os.path.join(os.getcwd(), f"temp_img_{i}.jpg")
                image.save(temp_path, format='JPEG')
                self.add_page()
                self.set_font("Arial", 'I', 10)
                self.cell(0, 10, f"Imagem {i+1}", ln=True, align='C')
                self.image(temp_path, x=30, w=150)
                os.remove(temp_path)
            except Exception as e:
                print(f"Erro ao adicionar imagem {i+1}: {e}")

    def gerar_pdf(self, nome_arquivo):
        os.makedirs(PASTA_SAIDA, exist_ok=True)
        caminho_pdf = os.path.join(PASTA_SAIDA, nome_arquivo)
        self.add_page()
        self.add_descricao()
        self.add_imagens()
        self.output(caminho_pdf)

def obter_access_token(token_input):
    if token_input.startswith("TG-"):
        data = {
            "grant_type": "authorization_code",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "code": token_input,
            "redirect_uri": REDIRECT_URI
        }
    elif len(token_input) > 80 and "-" in token_input:
        data = {
            "grant_type": "refresh_token",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "refresh_token": token_input
        }
    else:
        return token_input

    res = requests.post("https://api.mercadolibre.com/oauth/token", data=data)
    if res.status_code != 200:
        messagebox.showerror("Erro ao gerar access_token", res.text)
        return None

    return res.json().get("access_token")

def extrair_dados_anuncio(mlb_id, access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    url_item = f"https://api.mercadolibre.com/items/{mlb_id}"
    res_item = requests.get(url_item, headers=headers)
    if res_item.status_code != 200:
        print(f"[ERRO] Código {res_item.status_code} ao consultar {mlb_id}")
        return None, None, [], mlb_id

    item_data = res_item.json()
    print(f"[API] Status item: {res_item.status_code} - {mlb_id}")
    print("[API] Chaves recebidas:", list(item_data.keys()))

    titulo = item_data.get("title", "Sem título")
    pictures = item_data.get("pictures", [])
    if not pictures:
        print(f"[AVISO] Nenhuma imagem retornada para {mlb_id}")
    imagens = [img["secure_url"] for img in pictures if "secure_url" in img]

    url_desc = f"https://api.mercadolibre.com/items/{mlb_id}/description"
    res_desc = requests.get(url_desc, headers=headers)
    descricao_real = res_desc.json().get("plain_text", "") if res_desc.status_code == 200 else ""

    sku = item_data.get("seller_custom_field", None)
    descricao_completa = f"SKU: {sku}\n\n{descricao_real}" if sku else descricao_real or "Sem descrição encontrada."

    return titulo, descricao_completa, imagens, mlb_id

def selecionar_excel():
    global EXCEL_PATH
    EXCEL_PATH = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if EXCEL_PATH:
        lbl_excel.config(text=f"Planilha: {os.path.basename(EXCEL_PATH)}")

def escolher_pasta_saida():
    global PASTA_SAIDA
    folder = filedialog.askdirectory()
    if folder:
        PASTA_SAIDA = folder
        lbl_saida.config(text=f"Pasta de saída: {folder}")

def gerar():
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        messagebox.showerror("Erro", "Selecione uma planilha Excel válida.")
        return

    if not PASTA_SAIDA:
        messagebox.showerror("Erro", "Selecione uma pasta de saída.")
        return

    token = entrada_token.get()
    if not token:
        messagebox.showerror("Erro", "Informe o token de acesso.")
        return

    access_token = obter_access_token(token)
    if not access_token:
        return

    df = pd.read_excel(EXCEL_PATH)
    mlbs = df['MLB'].dropna().astype(str).tolist()

    if not mlbs:
        messagebox.showwarning("Aviso", "A planilha está vazia ou sem códigos MLB.")
        return

    barra["maximum"] = len(mlbs)
    for i, mlb_id in enumerate(mlbs):
        status_label.config(text=f"Gerando {mlb_id} ({i+1}/{len(mlbs)})")
        root.update()
        try:
            titulo, descricao, imagens, mlb_codigo = extrair_dados_anuncio(mlb_id, access_token)
            if imagens:
                pdf = AnuncioPDF(titulo, descricao, imagens)
                pdf.gerar_pdf(f"{mlb_codigo}.pdf")
                log_resultados.append(f"{mlb_id}: SUCESSO")
            else:
                log_resultados.append(f"{mlb_id}: SEM IMAGENS")
        except Exception as e:
            log_resultados.append(f"{mlb_id}: ERRO - {e}")
        barra["value"] = i + 1
        root.update()

    with open(os.path.join(PASTA_SAIDA, "relatorio.txt"), "w", encoding="utf-8") as f:
        for linha in log_resultados:
            f.write(linha + "\n")

    status_label.config(text="✅ Todos os PDFs foram gerados com sucesso!")
    abrir_btn.config(state="normal")

# Interface
root = tk.Tk()
root.title("Gerador de PDF Cyber")
root.geometry("500x380")
root.configure(bg="#f0f4f8")

tk.Label(root, text="Cole seu token:", bg="#f0f4f8", font=("Arial", 10, "bold")).pack(pady=(15,0))
entrada_token = tk.Entry(root, width=55)
entrada_token.pack(pady=(0,10))

tk.Button(root, text="Selecionar Planilha Excel", command=selecionar_excel, bg="#cce5ff").pack(pady=(5,0))
lbl_excel = tk.Label(root, text="Nenhuma planilha selecionada.", bg="#f0f4f8")
lbl_excel.pack(pady=(0,10))

tk.Button(root, text="Escolher Pasta de Saída", command=escolher_pasta_saida, bg="#cce5ff").pack(pady=(0,0))
lbl_saida = tk.Label(root, text="Nenhuma pasta de saída definida.", bg="#f0f4f8")
lbl_saida.pack(pady=(0,10))

tk.Button(root, text="Gerar PDFs", bg="#28a745", fg="white", font=("Arial", 10, "bold"), command=gerar).pack(pady=(15,5))

barra = ttk.Progressbar(root, length=400)
barra.pack(pady=(0,5))

status_label = tk.Label(root, text="", bg="#f0f4f8", font=("Arial", 9))
status_label.pack()

abrir_btn = tk.Button(root, text="Abrir pasta de saída", command=lambda: os.startfile(PASTA_SAIDA), state="disabled")
abrir_btn.pack(pady=(5,10))

root.mainloop()
