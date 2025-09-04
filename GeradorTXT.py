import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path


def gerar_txt():
    #seleciona o arquivo pelo windows explorer
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )

    if not caminho_arquivo:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
        return

    try:
        # Faz a leitura da planilha
        df = pd.read_excel(caminho_arquivo, dtype=str)
        df = df.fillna("").applymap(lambda x: str(x).strip())

        # monta as linhas já formatadas
        linhas_formatadas = []
        for _, row in df.iterrows():
            linha = (
                f"{int(row['Código']):06d},"
                f"{row['Tipo Cliente']},"
                f"\"{int(row['CNPJ/CPF']):014d}\","
                f"\"{row['Apelido']}\","
                f"\"{row['Nome']}\","
                f"\"{row['Endereço']}\","
                f"\"{row['Complemento']}\","
                f"\"{row['Bairro']}\","
                f"\"{row['CEP']}\","
                f"{row['Cod. Cidade']},"
                f"\"{row['Cidade']}\","
                f"\"{row['UF']}\","
                f"\"{row['Telefone']}\","
                f"\"{row['Fax']}\","
                f"\"{row['IE']}\","
                f"\"{row['Conta Contábil']}\","
                f"\"{row['Dt. Cadastro']}\","
                f"\"{row['IM']}\","
                f"\"{row['E-mail']}\","
                f"{row['País']},"
                f"\"{row['SUFRAMA']}\","
                f"\"{row['Reservado']}\","
                f"{row['Enq. Fed']}"
            )
            linhas_formatadas.append(linha)

        # Salva o arquivo TXT
        caminho_saida = Path(caminho_arquivo).with_suffix('.txt')
        with open(caminho_saida, 'w', encoding='utf-8') as f:
            f.write("\n".join(linhas_formatadas))

        messagebox.showinfo("Sucesso", f"Arquivo gerado em:\n{caminho_saida}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


# Criar janela principal
janela = tk.Tk()
janela.title("Excel → TXT")
janela.geometry("280x150")

tk.Label(janela, text="Gerar TXT participantes").pack(pady=10)
tk.Button(janela, text="Selecionar Planilha", command=gerar_txt).pack(pady=20)
tk.Button(janela, text="Sair", command=janela.quit).pack()

janela.mainloop()
