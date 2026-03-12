"""
Necessidade de produção de slitter
=======================================================
Regras de negócio:

01 - Exporta do sistema o cronograma de todas as máquinas
02 - Usando as OPs do cronograma exporta os itens das OPs
03 - Exporta o ZPP001 com o estoque 

Condições:
Usar a referência à coluna de estoque "Utilização livre" do relatório ZPP001
Informar no relatório final a coluna "Estq.Plan."

No software C# acrescentar a coluna programada que está em máquina.

Observação:
No relatório final de necessidade deixar as 3 colunas disponíveis mas uma com cálculo final de necessidade
"Qtd_pendente" (quantidade que ainda falta retirar do estoque)
"""

import pandas as pd
import os
import platform

SO = platform.system()

# =============================================================================
# 1. CARGA DOS ARQUIVOS
# =============================================================================

def get_current_user():
    if SO == 'Windows':
        return os.getenv('USERNAME')
    else:
        return os.getenv('USER')

USUARIO = get_current_user()
print(f"Usuário: {USUARIO}")

if SO == 'Windows':
    BASE_INPUT  = r'C:\Users\jefersson.souza\OneDrive - Açotel Indústria e Comércio LTDA\Dev\Necessidade_Slitter_py\Files\input'
    BASE_OUTPUT = r'C:\Users\jefersson.souza\OneDrive - Açotel Indústria e Comércio LTDA\Dev\Necessidade_Slitter_py\Files\output'
else:
    BASE_INPUT  = r'/home/stark/Documentos/Dev/Necessidade_Slitter_py/Files/input/'
    BASE_OUTPUT = r'/home/stark/Documentos/Dev/Necessidade_Slitter_py/Files/output'

print(f"Sistema Operacional: {SO}")

# Carregamento dos cronogramas (CR)
df_CR_itl50_01  = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL50-1-EXPORT.xlsx'))
df_CR_itl50_02  = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL50-2-EXPORT.xlsx'))
df_CR_itl75_01  = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL75-1-EXPORT.xlsx'))
df_CR_itl100_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL100-1-EXPORT.xlsx'))
df_CR_itl130_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL130-1-EXPORT.xlsx'))

df_CR_perf75_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-PERF75-1-EXPORT.xlsx'))
df_CR_perf75_02 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-PERF75-2-EXPORT.xlsx'))
df_CR_perf85_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-PERF85-1-EXPORT.xlsx'))

# Carregamento dos itens
df_itl50_01  = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL50-1-EXPORT.xlsx'))
df_itl50_02  = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL50-2-EXPORT.xlsx'))
df_itl75_01  = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL75-1-EXPORT.xlsx'))
df_itl100_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL100-1-EXPORT.xlsx'))
df_itl130_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL130-1-EXPORT.xlsx'))

df_perf75_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-PERF75-1-EXPORT.xlsx'))
df_perf75_02 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-PERF75-2-EXPORT.xlsx'))
df_perf85_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-PERF85-1-EXPORT.xlsx'))

# =============================================================================
# 2. PREPARAÇÃO DOS DADOS
# =============================================================================

# Concatenar todos os cronogramas
df_cronograma_grup = pd.concat(
    [df_CR_itl50_01, df_CR_itl50_02, df_CR_itl75_01, df_CR_itl100_01, df_CR_itl130_01, 
     df_CR_perf75_01, df_CR_perf85_01],
    ignore_index=True
)

# Concatenar todos os itens
df_itens_cronograma_grup = pd.concat(
    [df_itl50_01, df_itl50_02, df_itl75_01, df_itl100_01, df_itl130_01, 
     df_perf75_01, df_perf85_01],
    ignore_index=True
)

# FILTRO CRÍTICO: Eliminar materiais não produtivos (SUCATA e 2ª QUALIDADE) e qtd < 1
# Esses materiais aparecem com quantidade negativa ou zero e não representam necessidade real
df_itens_cronograma = df_itens_cronograma_grup[
    df_itens_cronograma_grup['Qtd.necessária (EINHEIT)'] >= 1
].copy()

# CÁLCULO DE Qtd_pendente (Quantidade que ainda falta retirar)
# Fórmula: Qtd_pendente = Qtd.necessária - Qtd.retirada
# Sua fórmula (Retirada - Necessária) * (-1) é algebricamente idêntica ✓
df_itens_cronograma['Qtd_pendente'] = (
    df_itens_cronograma['Qtd.necessária (EINHEIT)'] - 
    df_itens_cronograma['Qtd.retirada (EINHEIT)']
)

# IMPORTANTE: Qtd_pendente pode ser negativo se já foi retirado mais que o necessário.
# Nesse caso, não há necessidade de retirar mais material.
# Vamos garantir que valores negativos sejam tratados como zero na necessidade final
df_itens_cronograma['Qtd_pendente'] = df_itens_cronograma['Qtd_pendente'].clip(lower=0)

# Agrupar por Ordem + Material, somando as quantidades pendentes
df_itens_cronograma = (
    df_itens_cronograma
    .groupby(['Ordem', 'Material', 'Texto breve material'], as_index=False)['Qtd_pendente']
    .sum()
)

# Filtrar apenas itens com quantidade pendente > 0 (há algo a retirar)
df_itens_cronograma = df_itens_cronograma[df_itens_cronograma['Qtd_pendente'] > 0]

# =============================================================================
# 3. EXTRAÇÃO DE DATAS DO CRONOGRAMA
# =============================================================================

# FIX CRÍTICO: Extrair datas diretamente do CR bruto, sem filtrar por Quantidade.
# Algumas ordens (tipo ZPP2) têm quantidade negativa no CR, mas são válidas
# e têm datas de sequenciamento que precisam ser associadas aos itens.
df_datas_ordens = (
    df_cronograma_grup
    .groupby('Ordem', as_index=False)['Data sequenciamento']
    .min()  # Data mais cedo da ordem (critério FIFO)
)

# Fazer merge para trazer as datas para os itens
df_saldo_prod = df_itens_cronograma.merge(df_datas_ordens, on='Ordem', how='left')

# =============================================================================
# 4. PREPARAÇÃO DO ESTOQUE
# =============================================================================

# Carregar relatório de estoque
df_zpp001 = pd.read_excel(os.path.join(BASE_INPUT, 'ZPP001-EXPORT.xlsx'))

# Selecionar apenas as colunas necessárias
df_estoque = df_zpp001[[
    'Material', 
    'Utilização livre', 
    'Denom.grupo merc.',
    'Matriz de Conformação', 
    'Espessura Padrão (mm)'
]].copy()

# Converter 'Utilização livre' para numérico, tratando erros como 0
df_estoque['Utilização livre'] = pd.to_numeric(
    df_estoque['Utilização livre'], 
    errors='coerce'
).fillna(0)

# FIX: Manter TODOS os materiais FITA SLITTER, incluindo os com estoque zero.
# Isso é essencial para identificar necessidades de materiais sem estoque disponível.
df_estoque = df_estoque[
    df_estoque['Denom.grupo merc.'] == 'IN - FITA SLITTER'
].reset_index(drop=True)

# =============================================================================
# 5. MERGE — Trazer estoque para o cronograma
# =============================================================================

df_necessidade = df_saldo_prod.merge(
    df_estoque[['Material', 'Utilização livre', 'Matriz de Conformação', 'Espessura Padrão (mm)']],
    on='Material',
    how='left'
)

# Preencher NaN com 0 para materiais sem cadastro ou sem grupo FITA SLITTER
df_necessidade['Utilização livre'] = df_necessidade['Utilização livre'].fillna(0)

# =============================================================================
# 6. ORDENAÇÃO: Data sequenciamento (FIFO)
# =============================================================================

# Ordenar por data para aplicar lógica FIFO (First In, First Out)
df_necessidade = df_necessidade.sort_values('Data sequenciamento').reset_index(drop=True)

# =============================================================================
# 7. CÁLCULO FIFO — Saldo consumido por material em ordem cronológica
# =============================================================================

# Demanda acumulada: quanto já foi "comprometido" deste material nas ordens anteriores
df_necessidade['Demanda Acumulada'] = (
    df_necessidade.groupby('Material')['Qtd_pendente'].cumsum()
)

# Saldo projetado: quanto vai sobrar após atender esta ordem
# Se negativo, significa que não há estoque suficiente
df_necessidade['Saldo Projetado'] = (
    df_necessidade['Utilização livre'] - df_necessidade['Demanda Acumulada']
)

# Status: indica se o estoque atende ou se é preciso programar compra/produção
df_necessidade['Status'] = df_necessidade['Saldo Projetado'].apply(
    lambda x: 'Ok' if x >= 0 else 'Programar'
)

# =============================================================================
# 8. COLUNAS FINAIS
# =============================================================================

df_necessidade = df_necessidade[[
    'Ordem',
    'Material',
    'Texto breve material',
    'Espessura Padrão (mm)',
    'Matriz de Conformação',
    'Data sequenciamento',
    'Qtd_pendente',
    'Utilização livre',
    'Demanda Acumulada',
    'Saldo Projetado',
    'Status'
]]

# =============================================================================
# 9. EXPORTAÇÃO
# =============================================================================

os.makedirs(BASE_OUTPUT, exist_ok=True)
caminho_saida = os.path.join(BASE_OUTPUT, 'Necessidade - Slitter - Revisado.xlsx')
df_necessidade.to_excel(caminho_saida, index=False)

print(f"\nExportado: {caminho_saida}")
print(f"Total de linhas: {len(df_necessidade)}")
print(f"\nResumo de Status:")
print(df_necessidade['Status'].value_counts().to_string())

# Estatísticas adicionais
print(f"\nMateriais distintos: {df_necessidade['Material'].nunique()}")
print(f"Ordens distintas: {df_necessidade['Ordem'].nunique()}")
print(f"\nItens que precisam ser programados:")
programar = df_necessidade[df_necessidade['Status'] == 'Programar']
print(f"  Linhas: {len(programar)}")
print(f"  Materiais distintos: {programar['Material'].nunique()}")