import pandas as pd
import itertools
import re

# 1. Carregar as bases de candidatos do TSE (2020 e 2024)
# Certifique-se de que os arquivos CSV do TSE estejam no diretório de execução
df_2020 = pd.read_csv('consulta_cand_2020_SP.csv', encoding='latin1', sep=';')
df_2024 = pd.read_csv('consulta_cand_2024_SP.csv', encoding='latin1', sep=';')

MUNICIPIO = 'FRANCO DA ROCHA'

# 2. Filtrar município de Franco da Rocha E apenas o cargo de VEREADOR (CD_CARGO = 13)
cand_2020 = df_2020[
    (df_2020['NM_UE'].str.upper() == MUNICIPIO) & 
    ((df_2020['CD_CARGO'] == 13) | (df_2020['DS_CARGO'].str.upper() == 'VEREADOR'))
].copy()

cand_2024 = df_2024[
    (df_2024['NM_UE'].str.upper() == MUNICIPIO) & 
    ((df_2024['CD_CARGO'] == 13) | (df_2024['DS_CARGO'].str.upper() == 'VEREADOR'))
].copy()

# Coligações dos candidatos a prefeito em cada eleição
pref_2020 = df_2020[
    (df_2020['NM_UE'].str.upper() == MUNICIPIO) &
    (df_2020['DS_CARGO'].str.upper() == 'PREFEITO')
].copy()

pref_2024 = df_2024[
    (df_2024['NM_UE'].str.upper() == MUNICIPIO) &
    (df_2024['DS_CARGO'].str.upper() == 'PREFEITO')
].copy()

def mapear_partidos_por_coligacao(candidatos_prefeito, candidatos_vereador):
    partidos_vereador = set(candidatos_vereador['SG_PARTIDO'].dropna().str.strip())
    coligacao_por_partido = {}

    for id_coligacao, (_, prefeito) in enumerate(candidatos_prefeito.iterrows()):
        composicao = str(prefeito['DS_COMPOSICAO_COLIGACAO']).upper()
        for partido in partidos_vereador:
            padrao = rf'(?<![A-Z0-9]){re.escape(partido.upper())}(?![A-Z0-9])'
            if re.search(padrao, composicao):
                coligacao_por_partido[partido] = id_coligacao

    return coligacao_por_partido

coligacao_2020_por_partido = mapear_partidos_por_coligacao(pref_2020, cand_2020)
coligacao_2024_por_partido = mapear_partidos_por_coligacao(pref_2024, cand_2024)

# 3. Identificar os vereadores que disputaram AMBOS os pleitos (Painel Longitudinal)
titulos_2020 = set(cand_2020['NR_TITULO_ELEITORAL_CANDIDATO'])
titulos_2024 = set(cand_2024['NR_TITULO_ELEITORAL_CANDIDATO'])
titulos_comuns = list(titulos_2020.intersection(titulos_2024))

# Filtrar e remover duplicados mantendo unicidade pelo Título Eleitoral
cand_2020_p = cand_2020[cand_2020['NR_TITULO_ELEITORAL_CANDIDATO'].isin(titulos_comuns)].drop_duplicates(subset=['NR_TITULO_ELEITORAL_CANDIDATO'])
cand_2024_p = cand_2024[cand_2024['NR_TITULO_ELEITORAL_CANDIDATO'].isin(titulos_comuns)].drop_duplicates(subset=['NR_TITULO_ELEITORAL_CANDIDATO'])

# Mapeamentos: Título Eleitoral -> Nome / Coligação do candidato a prefeito apoiado
nome_dict = dict(zip(cand_2020_p['NR_TITULO_ELEITORAL_CANDIDATO'], cand_2020_p['NM_CANDIDATO']))

colig_2020 = dict(zip(
    cand_2020_p['NR_TITULO_ELEITORAL_CANDIDATO'],
    cand_2020_p['SG_PARTIDO'].map(coligacao_2020_por_partido)
))

colig_2024 = dict(zip(
    cand_2024_p['NR_TITULO_ELEITORAL_CANDIDATO'],
    cand_2024_p['SG_PARTIDO'].map(coligacao_2024_por_partido)
))

# 4. Gerar combinações pareadas entre os vereadores
pares = list(itertools.combinations(titulos_comuns, 2))

registros_2020 = []
registros_2024 = []
registros_longitudinal = []

def classificar_transicao(alianca_20, alianca_24):
    if not alianca_20 and not alianca_24:
        return 1  # Rival Estável
    elif not alianca_20 and alianca_24:
        return 2  # Virou Aliado
    elif alianca_20 and alianca_24:
        return 3  # Aliado Estável
    elif alianca_20 and not alianca_24:
        return 4  # Virou Rival

for c1, c2 in pares:
    nome_a = nome_dict[c1]
    nome_b = nome_dict[c2]
    
    # Avaliação das alianças
    a_20 = int(pd.notna(colig_2020[c1]) and colig_2020[c1] == colig_2020[c2])
    a_24 = int(pd.notna(colig_2024[c1]) and colig_2024[c1] == colig_2024[c2])
    
    cat_long = classificar_transicao(bool(a_20), bool(a_24))
    
    # Gravação simétrica para matrizes (A -> B e B -> A)
    for ca, cb in [(nome_a, nome_b), (nome_b, nome_a)]:
        registros_2020.append({'Cand_A': ca, 'Cand_B': cb, 'Valor': a_20})
        registros_2024.append({'Cand_A': ca, 'Cand_B': cb, 'Valor': a_24})
        registros_longitudinal.append({'Cand_A': ca, 'Cand_B': cb, 'Valor': cat_long})

# 5. Construção das Matrizes Pivotadas
matriz_2020 = pd.DataFrame(registros_2020).pivot(index='Cand_A', columns='Cand_B', values='Valor')
matriz_2024 = pd.DataFrame(registros_2024).pivot(index='Cand_A', columns='Cand_B', values='Valor')
matriz_long = pd.DataFrame(registros_longitudinal).pivot(index='Cand_A', columns='Cand_B', values='Valor')

# 6. Exportação para arquivo Excel multi-abas
caminho_saida = 'relacoes_vereadores_franco_da_rocha_coligacoes_prefeitos.xlsx'

with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
    matriz_2020.to_excel(writer, sheet_name='Eleicoes_2020')
    matriz_2024.to_excel(writer, sheet_name='Eleicoes_2024')
    matriz_long.to_excel(writer, sheet_name='Transicao_Longitudinal')

print(f"Arquivo Excel gerado com sucesso em: {caminho_saida}")
print(f"Total de vereadores repetidos analisados: {len(titulos_comuns)}")