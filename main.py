import pandas as pd
from deep_translator import GoogleTranslator
import time
from tkinter import Tk, filedialog


Tk().withdraw()
caminho_arquivo = filedialog.askopenfilename(
    title='Selecione a planilha Excel em chinês',
    filetypes=[('Planilhas Excel', '*.xlsx *.xls')]
)

if not caminho_arquivo:
    print("Nenhum arquivo selecionado.")
    exit()

print(f"Arquivo selecionado: {caminho_arquivo}")

df = pd.read_excel(caminho_arquivo)

def traduzir_celula(celula):
    if isinstance(celula, str) and celula.strip():
        try:
            return GoogleTranslator(source='auto', target='pt').translate(celula)
        except Exception as e:
            print(f"Erro ao traduzir: {e}")
            return celula
    return celula

colunas_traduzidas = []
for coluna in df.columns:
    try:
      col_traduzidas = GoogleTranslator(source='auto', target='pt').translate(coluna)
    except Exception as e:
      print(f"Erro ao traduzir o nome da coluna: {e}")
      col_traduzidas = coluna
    colunas_traduzidas.append(col_traduzidas)
df.columns = colunas_traduzidas


linhas_traduzidas = []
for i, linha in df.iterrows():
    linha_traduzida = linha.apply(traduzir_celula)
    linhas_traduzidas.append(linha_traduzida)

    if (i + 1) % 100 == 0:
        print(f"{i + 1} linhas traduzidas. Pausando 10 segundos para evitar bloqueio.")
        time.sleep(10)


df_traduzido = pd.DataFrame(linhas_traduzidas)
novo_nome = caminho_arquivo.replace('.xlsx', '_traduzido.xlsx')
df_traduzido.to_excel(novo_nome, index=False)

print(f"✅ Tradução concluída. Arquivo salvo como:\n{novo_nome}")
