import pandas as pd
import re
from tkinter import filedialog
from tkinter import *

# Cria a janela de seleção de arquivo
root = Tk()
root.filename = filedialog.askopenfilename(
    initialdir="/", title="Selecione o arquivo", filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
root.withdraw()

# Lê o arquivo Excel
df = pd.read_excel(root.filename)

# Função para corrigir DDD1 e número de FONE1


def corrigir_FONE1(DDD1, FONE1):
    if pd.isna(DDD1) or pd.isna(FONE1):
        return None

    # Limpa o número de FONE1
    FONE1 = re.sub(r'\D', '', str(FONE1))

    # Verifica se o número de FONE1 tem o tamanho correto
    if len(FONE1) != 8 and len(FONE1) != 9:
        return None

    return '{} {}'.format(DDD1, FONE1)

# Função para corrigir endereço


# Aplica as correções aos dados
df['FONE1'] = df.apply(lambda x: corrigir_FONE1(
    x['DDD1'], x['FONE1']), axis=1)

# Salva o arquivo corrigido
df.to_excel('corrigido.xlsx', index=False)
