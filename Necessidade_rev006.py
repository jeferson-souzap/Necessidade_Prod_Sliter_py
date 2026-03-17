"""
Necessidade de produção de slitter
=======================================================
Regras de negócio:

01 - Exporta do sistema o cronograma de todas as máquinas
02 - Usando as OPs do cronograma exporta os itens das OPs
03 - Exporta o ZPP001 com o estoque
04 - Exporta o PROG (ZPP1) com as ordens de produção de fita programadas

Condições:
- Usar a referência à coluna de estoque "Utilização livre" do relatório ZPP001
- Considerar também as quantidades em produção (PROG) que ainda não estão no estoque físico
- O saldo projetado é calculado sobre: Utilização livre + Qtd_Programada (PROG líquido)
- Informar no relatório final as colunas "Utilização livre", "Qtd_Programada" e "Saldo Inicial"

rev005 — Novidade:
    Acrescentada a consideração das OPs do PROG (ZPP1) no cálculo do saldo.
    Qtd_Programada = soma de (Qtd da ordem - Qtd já fornecida) por material no PROG.
    Saldo Inicial  = Utilização livre + Qtd_Programada
    O FIFO agora consome a partir do Saldo Inicial, não apenas do estoque físico.

rev006 — Novidade:
    Validação de arquivos obrigatórios antes da carga.
    O script lista todos os arquivos ausentes de uma vez e interrompe a execução
    com mensagem clara antes de tentar qualquer leitura.
"""

from datetime import datetime
import pandas as pd
import os
import platform
import sys

SO = platform.system()

# =============================================================================
# 1. PATHS
# =============================================================================

def get_current_user():
    return os.getenv('USERNAME') if SO == 'Windows' else os.getenv('USER')

USUARIO = get_current_user()
print(f"Usuário: {USUARIO} | SO: {SO}")

if SO == 'Windows':
    BASE_INPUT  = r'D:\#Mega\Jeferson - Dev\02 - Linguagens\Python\Acotel\Necessidade_Prod_Sliter_py\Files\input'
    BASE_OUTPUT = r'D:\#Mega\Jeferson - Dev\02 - Linguagens\Python\Acotel\Necessidade_Prod_Sliter_py\Files\output'
else:
    BASE_INPUT  = r'/home/stark/Documentos/Dev/Necessidade_Slitter_py/Files/input/'
    BASE_OUTPUT = r'/home/stark/Documentos/Dev/Necessidade_Slitter_py/Files/output'

# =============================================================================
# 2. VALIDAÇÃO DOS ARQUIVOS
# =============================================================================

ARQUIVOS_OBRIGATORIOS = [
    # Cronogramas
    'CR-ITL50-1-EXPORT.xlsx',
    'CR-ITL50-2-EXPORT.xlsx',
    'CR-ITL75-1-EXPORT.xlsx',
    'CR-ITL100-1-EXPORT.xlsx',
    'CR-ITL130-1-EXPORT.xlsx',
    'CR-PERF75-1-EXPORT.xlsx',
    'CR-PERF85-1-EXPORT.xlsx',
    # Itens
    'ITENS-ITL50-1-EXPORT.xlsx',
    'ITENS-ITL50-2-EXPORT.xlsx',
    'ITENS-ITL75-1-EXPORT.xlsx',
    'ITENS-ITL100-1-EXPORT.xlsx',
    'ITENS-ITL130-1-EXPORT.xlsx',
    'ITENS-PERF75-1-EXPORT.xlsx',
    'ITENS-PERF85-1-EXPORT.xlsx',
    # Estoque e programação
    'ZPP001-EXPORT.xlsx',
    'PROG-EXPORT.xlsx',
]

arquivos_ausentes = [
    nome for nome in ARQUIVOS_OBRIGATORIOS
    if not os.path.isfile(os.path.join(BASE_INPUT, nome))
]

if arquivos_ausentes:
    print("\n" + "=" * 60)
    print("  ERRO — Arquivos obrigatorios nao encontrados:")
    print("=" * 60)
    for nome in arquivos_ausentes:
        print(f"  [AUSENTE]  {nome}")
    print("=" * 60)
    print(f"\n  Pasta esperada: {BASE_INPUT}")
    print(f"  Total de arquivos faltando: {len(arquivos_ausentes)}\n")
    sys.exit(1)

print("Todos os arquivos encontrados. Iniciando processamento...\n")

# =============================================================================
# 3. CARGA DOS ARQUIVOS
# =============================================================================

# Cronogramas (CR)
df_CR_itl50_01  = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL50-1-EXPORT.xlsx'))
df_CR_itl50_02  = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL50-2-EXPORT.xlsx'))
df_CR_itl75_01  = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL75-1-EXPORT.xlsx'))
df_CR_itl100_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL100-1-EXPORT.xlsx'))
df_CR_itl130_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-ITL130-1-EXPORT.xlsx'))
df_CR_perf75_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-PERF75-1-EXPORT.xlsx'))
df_CR_perf85_01 = pd.read_excel(os.path.join(BASE_INPUT, 'CR-PERF85-1-EXPORT.xlsx'))

# Itens das ordens
df_itl50_01  = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL50-1-EXPORT.xlsx'))
df_itl50_02  = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL50-2-EXPORT.xlsx'))
df_itl75_01  = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL75-1-EXPORT.xlsx'))
df_itl100_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL100-1-EXPORT.xlsx'))
df_itl130_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-ITL130-1-EXPORT.xlsx'))
df_perf75_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-PERF75-1-EXPORT.xlsx'))
df_perf85_01 = pd.read_excel(os.path.join(BASE_INPUT, 'ITENS-PERF85-1-EXPORT.xlsx'))

# Estoque e programação
df_zpp001 = pd.read_excel(os.path.join(BASE_INPUT, 'ZPP001-EXPORT.xlsx'))
df_prog   = pd.read_excel(os.path.join(BASE_INPUT, 'PROG-EXPORT.xlsx'))

# =============================================================================
# 4. PREPARAÇÃO DO CRONOGRAMA E ITENS
# =============================================================================

df_cronograma_grup = pd.concat(
    [df_CR_itl50_01, df_CR_itl50_02, df_CR_itl75_01, df_CR_itl100_01, df_CR_itl130_01,
     df_CR_perf75_01, df_CR_perf85_01],
    ignore_index=True
)

df_itens_cronograma_grup = pd.concat(
    [df_itl50_01, df_itl50_02, df_itl75_01, df_itl100_01, df_itl130_01,
     df_perf75_01, df_perf85_01],
    ignore_index=True
)

# Eliminar materiais não produtivos (SUCATA, 2ª QUALIDADE) com qtd < 1
df_itens_cronograma = df_itens_cronograma_grup[
    df_itens_cronograma_grup['Qtd.necessária (EINHEIT)'] >= 1
].copy()

# Quantidade pendente: o que ainda falta retirar do estoque
df_itens_cronograma['Qtd_pendente'] = (
    df_itens_cronograma['Qtd.necessária (EINHEIT)'] -
    df_itens_cronograma['Qtd.retirada (EINHEIT)']
).clip(lower=0)

# Agrupar por Ordem + Material
df_itens_cronograma = (
    df_itens_cronograma
    .groupby(['Ordem', 'Material', 'Texto breve material'], as_index=False)['Qtd_pendente']
    .sum()
)

# Manter apenas itens com pendência real
df_itens_cronograma = df_itens_cronograma[df_itens_cronograma['Qtd_pendente'] > 0]

# =============================================================================
# 5. DATAS DO CRONOGRAMA (FIFO)
# =============================================================================

df_datas_ordens = (
    df_cronograma_grup
    .groupby('Ordem', as_index=False)['Data sequenciamento']
    .min()
)

df_saldo_prod = df_itens_cronograma.merge(df_datas_ordens, on='Ordem', how='left')

# =============================================================================
# 6. ESTOQUE FÍSICO (ZPP001)
# =============================================================================

df_estoque = df_zpp001[[
    'Material',
    'Utilização livre',
    'Denom.grupo merc.',
    'Matriz de Conformação',
    'Espessura Padrão (mm)'
]].copy()

df_estoque['Utilização livre'] = pd.to_numeric(
    df_estoque['Utilização livre'], errors='coerce'
).fillna(0)

# Manter todos os materiais FITA SLITTER (inclusive os com estoque zero)
df_estoque = df_estoque[
    df_estoque['Denom.grupo merc.'] == 'IN - FITA SLITTER'
].reset_index(drop=True)

# =============================================================================
# 7. ORDENS PROGRAMADAS (PROG / ZPP1) — NOVIDADE rev005
# =============================================================================
# O PROG contém Ordens de Produção (ZPP1) para fabricar fita slitter.
# Essas quantidades ainda não estão no estoque físico ("Utilização livre"),
# mas já estão comprometidas para produção e devem reduzir a necessidade real.
#
# Qtd_Programada = quantidade total da ordem - quantidade já fornecida/entregue
# (descontar o fornecido evita dupla contagem com o estoque físico)

df_prog['Qtd_Programada'] = (
    pd.to_numeric(df_prog['Quantidade da ordem (GMEIN)'], errors='coerce').fillna(0) -
    pd.to_numeric(df_prog['Qtd.fornecida (GMEIN)'], errors='coerce').fillna(0)
).clip(lower=0)

df_prog_agg = (
    df_prog
    .groupby('Material', as_index=False)['Qtd_Programada']
    .sum()
)

# =============================================================================
# 8. MERGE — Estoque + Programação + Cronograma
# =============================================================================

df_necessidade = df_saldo_prod.merge(
    df_estoque[['Material', 'Utilização livre', 'Matriz de Conformação', 'Espessura Padrão (mm)']],
    on='Material',
    how='left'
)

# Trazer quantidades programadas do PROG
df_necessidade = df_necessidade.merge(df_prog_agg, on='Material', how='left')

# Preencher NaN
df_necessidade['Utilização livre'] = df_necessidade['Utilização livre'].fillna(0)
df_necessidade['Qtd_Programada']   = df_necessidade['Qtd_Programada'].fillna(0)

# Saldo Inicial = estoque físico + o que está sendo produzido (ainda não entregue)
df_necessidade['Saldo Inicial'] = (
    df_necessidade['Utilização livre'] + df_necessidade['Qtd_Programada']
)

# =============================================================================
# 9. ORDENAÇÃO E FIFO
# =============================================================================

df_necessidade = df_necessidade.sort_values('Data sequenciamento').reset_index(drop=True)

# Demanda acumulada por material em ordem cronológica
df_necessidade['Demanda Acumulada'] = (
    df_necessidade.groupby('Material')['Qtd_pendente'].cumsum()
)

# Saldo projetado: quanto sobra após atender cada ordem (usando o Saldo Inicial)
df_necessidade['Saldo Projetado'] = (
    df_necessidade['Saldo Inicial'] - df_necessidade['Demanda Acumulada']
)

# Status final
df_necessidade['Status'] = df_necessidade['Saldo Projetado'].apply(
    lambda x: 'Ok' if x >= 0 else 'Programar'
)

# =============================================================================
# 10. COLUNAS FINAIS
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
    'Qtd_Programada',
    'Saldo Inicial',
    'Demanda Acumulada',
    'Saldo Projetado',
    'Status'
]]

# =============================================================================
# 11. EXPORTAÇÃO
# =============================================================================

os.makedirs(BASE_OUTPUT, exist_ok=True)


timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

caminho_saida = os.path.join(BASE_OUTPUT, f'Necessidade_Slitter {timestamp}.xlsx')
df_necessidade.to_excel(caminho_saida, index=False)

print(f"\nExportado: {caminho_saida}")
print(f"Total de linhas: {len(df_necessidade)}")

print(f"\nResumo de Status:")
print(df_necessidade['Status'].value_counts().to_string())

print(f"\nMateriais distintos : {df_necessidade['Material'].nunique()}")
print(f"Ordens distintas    : {df_necessidade['Ordem'].nunique()}")

programar = df_necessidade[df_necessidade['Status'] == 'Programar']
print(f"\nItens que precisam ser programados:")
print(f"  Linhas              : {len(programar)}")
print(f"  Materiais distintos : {programar['Material'].nunique()}")

# Comparativo: impacto do PROG no resultado
print("\n=== Impacto das OPs Programadas (PROG) ===")
df_prog_impacto = df_necessidade[df_necessidade['Qtd_Programada'] > 0][
    ['Material', 'Texto breve material', 'Utilização livre', 'Qtd_Programada', 'Saldo Inicial']
].drop_duplicates('Material').sort_values('Qtd_Programada', ascending=False)
print(df_prog_impacto.to_string(index=False))