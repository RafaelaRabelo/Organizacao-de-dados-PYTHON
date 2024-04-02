# ENCERRADAS NA SEMANA
# SCRIPT - RAFAELA BERNARDES RABELO
# ORGANIZAÇÃO E GERAÇÃO DE DOCUMENTO PARA RELATÓRIO SEMANAL DE ACOMPANHAMENTO DE EMISSÕES DE FIIS
# BIBLIOTECA PARA INSTALAR: pip install pandas datetime, timedelta
import pandas as pd
from datetime import datetime, timedelta

# Caminho do arquivo e nome da aba
excel_file = r'C:\caminho\para\seu\arquivo.xlsx
sheet_name = 'Base_Dados'

# Ler toda a aba do Excel e informar que os cabeçalhos estão na segunda linha
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=1)

# Obter a data de hoje
hoje = datetime.now()

# Calcular a data do "dia 1" (segunda-feira) da semana passada
dia_1_semana_passada = hoje - timedelta(days=(hoje.weekday() - 0 + 8) % 7 + 7)

# Calcular a data da última sexta-feira (dia 5) da semana passada
sexta_passada = hoje - timedelta(days=(hoje.weekday() - 4 + 6) % 7)

# Filtrar as linhas com base na "Data Anúncio de Início"
emissoes_semana_anterior = df[(df['Data Encerrmanto'] >= dia_1_semana_passada) & (df['Data Encerrmanto'] <= sexta_passada)]

# Selecionar as colunas desejadas
colunas_desejadas = ['Fundo ', 'Segmento', 'Gestor', 'Administrador', 'Coordenador', 'Emissão',
                     'Taxa de distribuição por cota', 'Captação-Base', 'Captação final (R$ Milhões)',
                     '% Captado/Montante Inicial', 'Preço', 'Preço Tela (R$)', 'Preço Patrimonial (R$)']

# Filtrar as colunas desejadas e criar uma cópia
tabela_resultante = emissoes_semana_anterior[colunas_desejadas].copy()

# Organizar tipos e formatos das colunas
tabela_resultante['Taxa de distribuição por cota'] = tabela_resultante['Taxa de distribuição por cota'].apply(lambda x: f'{x * 10**-1:.2%}')
tabela_resultante['Captação-Base'] = tabela_resultante['Captação-Base'].apply(lambda x: f'R$ {x * 10**-6:.2f}')
tabela_resultante['Captação final (R$ Milhões)'] = tabela_resultante['Captação final (R$ Milhões)'].apply(lambda x: f'R$ {x * 10**-6:.2f}')
tabela_resultante['% Captado/Montante Inicial'] = tabela_resultante['% Captado/Montante Inicial'].apply(lambda x: f'{x:.2%}')
tabela_resultante['Preço'] = tabela_resultante['Preço'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Tela (R$)'] = tabela_resultante['Preço Tela (R$)'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Patrimonial (R$)'] = tabela_resultante['Preço Patrimonial (R$)'].apply(lambda x: f'R$ {x:.2f}')

# Converter 'Captação-Base' e 'Captação final (R$ Milhões)' para formato numérico
tabela_resultante['Captação-Base'] = pd.to_numeric(tabela_resultante['Captação-Base'].apply(lambda x: x[2:].replace(',', '')))
tabela_resultante['Captação final (R$ Milhões)'] = pd.to_numeric(tabela_resultante['Captação final (R$ Milhões)'].apply(lambda x: x[2:].replace(',', '')), errors='coerce')  # Adicionado 'errors' para lidar com valores não numéricos

# Organizar a tabela pela coluna 'Captação-Base' do maior para o menor
tabela_resultante = tabela_resultante.sort_values(by='Captação final (R$ Milhões)', key=pd.to_numeric, na_position='last', ascending=False)

# Adicionar a nova linha "Total" com as somas
nova_linha_total = {'Fundo ': '', 'Segmento': '', 'Gestor': '', 'Administrador': '',
                    'Coordenador': '', 'Emissão': '', 'Taxa de distribuição por cota': 'Total',
                    'Captação-Base': tabela_resultante["Captação-Base"].sum(),
                    'Captação final (R$ Milhões)': tabela_resultante["Captação final (R$ Milhões)"].sum(),
                    '% Captado/Montante Inicial': '',
                    'Preço': '', 'Preço Tela (R$)': '', 'Preço Patrimonial (R$)': ''}

# Transformar a nova linha "Total" em um DataFrame
nova_linha_total_df = pd.DataFrame([nova_linha_total])

# Concatenar o DataFrame resultante com a nova linha "Total"
tabela_resultante = pd.concat([tabela_resultante, nova_linha_total_df], ignore_index=True)

# Exibir a tabela resultante
tabela_resultante



# INICIADAS NA SEMANA
# SCRIPT - RAFAELA BERNARDES RABELO
# ORGANIZAÇÃO E GERAÇÃO DE DOCUMENTO PARA RELATÓRIO SEMANAL DE ACOMPANHAMENTO DE EMISSÕES DE FIIS
# BIBLIOTECA PARA INSTALAR: pip install pandas datetime, timedelta
import pandas as pd
from datetime import datetime, timedelta

# Caminho do arquivo e nome da aba
excel_file = r'C:\caminho\para\seu\arquivo.xlsx
sheet_name = 'Base_Dados'

# Ler toda a aba do Excel e informar que os cabeçalhos estão na segunda linha
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=1)

# Obter a data de hoje
hoje = datetime.now()

# Calcular a data do "dia 1" (segunda-feira) da semana passada
dia_1_semana_passada = hoje - timedelta(days=(hoje.weekday() - 0 + 8) % 7 + 7)

# Calcular a data da última sexta-feira (dia 5) da semana passada
sexta_passada = hoje - timedelta(days=(hoje.weekday() - 4 + 6) % 7)

# Filtrar as linhas com base na "Data Anúncio de Início"
emissoes_semana_anterior = df[(df['Data Anúncio de Início'] >= dia_1_semana_passada) & (df['Data Anúncio de Início'] <= sexta_passada)]

# Selecionar as colunas desejadas
colunas_desejadas = ['Fundo ', 'Segmento', 'Gestor', 'Administrador', 'Coordenador', 'Emissão',
                     'Taxa de distribuição por cota', 'Captação-Base', 'Preço', 'Preço Tela (R$)', 'Preço Patrimonial (R$)']

# Filtrar as colunas desejadas e criar uma cópia
tabela_resultante = emissoes_semana_anterior[colunas_desejadas].copy()

# Organizar tipos e formatos das colunas
tabela_resultante['Taxa de distribuição por cota'] = tabela_resultante['Taxa de distribuição por cota'].apply(lambda x: f'{x * 10**-2:.2%}')
tabela_resultante['Captação-Base'] = tabela_resultante['Captação-Base'].apply(lambda x: f'R$ {x*10**-6:.2f}')
tabela_resultante['Preço'] = tabela_resultante['Preço'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Tela (R$)'] = tabela_resultante['Preço Tela (R$)'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Patrimonial (R$)'] = tabela_resultante['Preço Patrimonial (R$)'].apply(lambda x: f'R$ {x:.2f}')

# Converter 'Captação-Base' para formato numérico
tabela_resultante['Captação-Base'] = tabela_resultante['Captação-Base'].apply(lambda x: pd.to_numeric(x[2:].replace(',', ''), errors='coerce'))

# Calcular a soma das colunas "Captação-Base"
soma_captação_base = tabela_resultante['Captação-Base'].sum()

# Organizar a tabela pela coluna 'Captação-Base' do maior para o menor
tabela_resultante = tabela_resultante.sort_values(by='Captação-Base', ascending=False)

# Adicionar a nova linha "Total" com as somas no final
nova_linha_total = {'Fundo ': '', 'Segmento': '', 'Gestor': '', 'Administrador': '',
                    'Coordenador': '', 'Emissão': '', 'Taxa de distribuição por cota': 'Total',
                    'Captação-Base': f'R$ {soma_captação_base:.2f}',
                    'Preço': '', 'Preço Tela (R$)': '', 'Preço Patrimonial (R$)': ''}

# Transformar a nova linha "Total" em um DataFrame
nova_linha_total_df = pd.DataFrame([nova_linha_total])

# Concatenar o DataFrame resultante com a nova linha "Total" no final
tabela_resultante = pd.concat([tabela_resultante, nova_linha_total_df], ignore_index=True)

# Substituir 'R$ nan' ou 'R$ 0.00' por espaços vazios nas colunas Preço Tela (R$) e Preço Patrimonial (R$)
colunas_substituir_nulos = ['Preço Tela (R$)', 'Preço Patrimonial (R$)']
tabela_resultante[colunas_substituir_nulos] = tabela_resultante[colunas_substituir_nulos].replace({'R$ nan': 'R$ -', 'R$ 0.00': 'R$ -'})

tabela_resultante





# INICIADAS NO MÊS
# SCRIPT - RAFAELA BERNARDES RABELO
# ORGANIZAÇÃO E GERAÇÃO DE DOCUMENTO PARA RELATÓRIO SEMANAL DE ACOMPANHAMENTO DE EMISSÕES DE FIIS
# BIBLIOTECA PARA INSTALAR: pip install pandas
import pandas as pd

# Caminho do arquivo e nome da aba
excel_file = r'C:\caminho\para\seu\arquivo.xlsx
sheet_name = 'Base_Dados'

# Ler toda a aba do Excel e informar que os cabeçalhos estão na segunda linha
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=1)

# Filtrar emissões encerradas no mês específico (por exemplo, janeiro de 2024)
mes_filtrado = '01/01/2024'
emissoes_iniciadas = df[df['Mês Início'].dt.strftime('%m/%d/%Y') == mes_filtrado]

# Selecionar as colunas desejadas
colunas_desejadas = ['Fundo ', 'Segmento', 'Gestor', 'Administrador', 'Coordenador', 'Emissão',
                     'Taxa de distribuição por cota', 'Captação-Base', 'Preço', 'Preço Tela (R$)', 'Preço Patrimonial (R$)']

# Filtrar as colunas desejadas e criar uma cópia
tabela_resultante = emissoes_iniciadas[colunas_desejadas].copy()

# Organizar tipos e formatos das colunas
tabela_resultante['Taxa de distribuição por cota'] = tabela_resultante['Taxa de distribuição por cota'].apply(lambda x: f'{x * 10**-2:.2%}')
tabela_resultante['Captação-Base'] = tabela_resultante['Captação-Base'].apply(lambda x: f'R$ {x*10**-6:.2f}')
tabela_resultante['Preço'] = tabela_resultante['Preço'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Tela (R$)'] = tabela_resultante['Preço Tela (R$)'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Patrimonial (R$)'] = tabela_resultante['Preço Patrimonial (R$)'].apply(lambda x: f'R$ {x:.2f}')

# Calcular a soma das colunas "Captação-Base"
soma_captação_base = tabela_resultante['Captação-Base'].apply(lambda x: float(x[2:].replace(',', ''))).sum()

# Converter 'Captação-Base' para formato numérico
tabela_resultante['Captação-Base'] = pd.to_numeric(tabela_resultante['Captação-Base'].apply(lambda x: x[2:].replace(',', '')))

# Organizar a tabela pela coluna 'Captação-Base' do maior para o menor, excluindo a última linha
tabela_resultante = tabela_resultante.iloc[:-1].sort_values(by='Captação-Base', ascending=False)

# Adicionar a nova linha "Total" com as somas no final
nova_linha_total = {'Fundo ': '', 'Segmento': '', 'Gestor': '', 'Administrador': '',
                    'Coordenador': '', 'Emissão': '', 'Taxa de distribuição por cota': 'Total',
                    'Captação-Base': f'R$ {soma_captação_base:.2f}',
                    'Preço': '', 'Preço Tela (R$)': '', 'Preço Patrimonial (R$)': ''}

# Transformar a nova linha "Total" em um DataFrame
nova_linha_total_df = pd.DataFrame([nova_linha_total])

# Concatenar o DataFrame resultante com a nova linha "Total" no final
tabela_resultante = pd.concat([tabela_resultante, nova_linha_total_df], ignore_index=True)

# Substituir 'R$ nan' ou 'R$ 0.00' por espaços vazios nas colunas Preço Tela (R$) e Preço Patrimonial (R$)
colunas_substituir_nulos = ['Preço Tela (R$)', 'Preço Patrimonial (R$)']
tabela_resultante[colunas_substituir_nulos] = tabela_resultante[colunas_substituir_nulos].replace({'R$ nan': 'R$ -', 'R$ 0.00': 'R$ -'})

# Exibir a tabela resultante
tabela_resultante




# ENCERRADAS NO MÊS
# SCRIPT - RAFAELA BERNARDES RABELO
# ORGANIZAÇÃO E GERAÇÃO DE DOCUMENTO PARA RELATÓRIO SEMANAL DE ACOMPANHAMENTO DE EMISSÕES DE FIIS
# BIBLIOTECA PARA INSTALAR: pip install pandas
import pandas as pd

# Caminho do arquivo e nome da aba
excel_file = r'C:\caminho\para\seu\arquivo.xlsx
sheet_name = 'Base_Dados'

# Ler toda a aba do Excel e informar que os cabeçalhos estão na segunda linha
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=1)

# Filtrar emissões encerradas no mês específico (por exemplo, janeiro de 2024)
mes_filtrado = '01/01/2024'
emissoes_encerradas = df[df['Mês Encerramento'].dt.strftime('%m/%d/%Y') == mes_filtrado]

# Selecionar as colunas desejadas
colunas_desejadas = ['Fundo ', 'Segmento', 'Gestor', 'Administrador', 'Coordenador', 'Emissão',
                     'Taxa de distribuição por cota', 'Captação-Base', 'Captação final (R$ Milhões)',
                     '% Captado/Montante Inicial', 'Preço', 'Preço Tela (R$)', 'Preço Patrimonial (R$)']

# Filtrar as colunas desejadas e criar uma cópia
tabela_resultante = emissoes_encerradas[colunas_desejadas].copy()

# Organizar tipos e formatos das colunas
tabela_resultante['Taxa de distribuição por cota'] = tabela_resultante['Taxa de distribuição por cota'].apply(lambda x: f'{x * 10**-1:.2%}')
tabela_resultante['Captação-Base'] = tabela_resultante['Captação-Base'].apply(lambda x: f'R$ {x * 10**-6:.2f}')
tabela_resultante['Captação final (R$ Milhões)'] = tabela_resultante['Captação final (R$ Milhões)'].apply(lambda x: f'R$ {x * 10**-6:.2f}')
tabela_resultante['% Captado/Montante Inicial'] = tabela_resultante['% Captado/Montante Inicial'].apply(lambda x: f'{x:.2%}')
tabela_resultante['Preço'] = tabela_resultante['Preço'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Tela (R$)'] = tabela_resultante['Preço Tela (R$)'].apply(lambda x: f'R$ {x:.2f}')
tabela_resultante['Preço Patrimonial (R$)'] = tabela_resultante['Preço Patrimonial (R$)'].apply(lambda x: f'R$ {x:.2f}')

# Converter 'Captação-Base' e 'Captação final (R$ Milhões)' para formato numérico
tabela_resultante['Captação-Base'] = pd.to_numeric(tabela_resultante['Captação-Base'].apply(lambda x: x[2:].replace(',', '')))
tabela_resultante['Captação final (R$ Milhões)'] = pd.to_numeric(tabela_resultante['Captação final (R$ Milhões)'].apply(lambda x: x[2:].replace(',', '')), errors='coerce')  # Adicionado 'errors' para lidar com valores não numéricos

# Organizar a tabela pela coluna 'Captação-Base' do maior para o menor
tabela_resultante = tabela_resultante.sort_values(by='Captação final (R$ Milhões)', key=pd.to_numeric, na_position='last', ascending=False)

# Adicionar a nova linha "Total" com as somas
nova_linha_total = {'Fundo ': '', 'Segmento': '', 'Gestor': '', 'Administrador': '',
                    'Coordenador': '', 'Emissão': '', 'Taxa de distribuição por cota': 'Total',
                    'Captação-Base': f'R$ {tabela_resultante["Captação-Base"].sum():.2f}',
                    'Captação final (R$ Milhões)': f'R$ {tabela_resultante["Captação final (R$ Milhões)"].sum():.2f}',
                    '% Captado/Montante Inicial': '',
                    'Preço': '', 'Preço Tela (R$)': '', 'Preço Patrimonial (R$)': ''}

# Transformar a nova linha "Total" em um DataFrame
nova_linha_total_df = pd.DataFrame([nova_linha_total])

# Concatenar o DataFrame resultante com a nova linha "Total"
tabela_resultante = pd.concat([tabela_resultante, nova_linha_total_df], ignore_index=True)

# Exibir a tabela resultante
tabela_resultante


