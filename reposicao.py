import pyodbc
import pandas as pd
import os
import openpyxl
import time

user_diretorio = os.path.expanduser('~')

user_diretorio = user_diretorio.replace("\\",'/')

servidor = 'numero_servidor'
dados = 'base_dados'
usuario = 'usuario_base'
senha = 'senha_base'

query = """
Query do SQLServer para puxar os dados do estoque da empresa. (Por motivos óbvios não colocarei a query aqui.)
"""

def read_SQL(server, database, username, password, query):
    connection = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
    conn = pyodbc.connect(connection)

    df = pd.read_sql_query(query, conn)

    df['filial'] = df['filial'].astype('string')
    df['produto'] = df['produto'].astype('string')
    df['COR_PRODUTO'] = df['COR_PRODUTO'].astype('string')
    df['GRADE'] = df['GRADE'].astype('string')

    conn.close()

    print('Dados do banco extraídos com sucesso!!')

    return df

def Reposicao(row):
    grade = {'XPP': ['34-35','XPP','U'],
            'PP': ['36-37', '38', 'PP'],
            'P': ['38-39', '40', 'P'],
            'M': ['40-41', '42', 'M'],
            'G': ['42-43', '44', 'G'],
            'GG': ['44-45', '46', 'GG']}

    if row['Bloqueado'] == 0:
        for chave, tam in grade.items():
            df_filtrado = df[(df['filial'] == row['filial']) & (df['produto'] == row['PRODUTO']) & (df['COR_PRODUTO'] == row['cor_produto']) & (df['GRADE'].isin(tam))]

            if not df_filtrado.empty:
                if df_filtrado['estoque_ideal_novo'].iloc[0] >= row[f'Estoque Mínimo {chave}'] and df_filtrado['estoque_ideal_novo'].iloc[0] <= row[f'Estoque Máximo {chave}']:
                    if row['Multiplo aposta'] > 0:
                        estoque_filial = df_filtrado['estoque_filial'].iloc[0] 
                        estoque_ideal_programado = round(df_filtrado['estoque_ideal_novo'].iloc[0] * row['Multiplo aposta'])
                        lojas.append(row['filial'])
                        produto.append(row['PRODUTO'])
                        desc_produto.append(row['DESC_PRODUTO'])
                        cor_produto.append(row['cor_produto'])
                        desc_cor_produto.append(df_filtrado['desc_cor_produto'].iloc[0])
                        colecao.append(row['colecao'])
                        aposta.append(row['Multiplo aposta'])
                        bloqueado.append(row['Bloqueado'])
                        tamanho.append(df_filtrado['GRADE'].iloc[0])
                        lista_estoque.append(estoque_ideal_programado)
                        reposicao.append(estoque_ideal_programado if (estoque_ideal_programado - estoque_filial) <= 0 else estoque_ideal_programado - estoque_filial)
                    else:
                        estoque_filial = df_filtrado['estoque_filial'].iloc[0] 
                        estoque_ideal_programado = df_filtrado['estoque_ideal_novo'].iloc[0] 
                        lojas.append(row['filial'])
                        produto.append(row['PRODUTO'])
                        desc_produto.append(row['DESC_PRODUTO'])
                        cor_produto.append(row['cor_produto'])
                        desc_cor_produto.append(df_filtrado['desc_cor_produto'].iloc[0])
                        colecao.append(row['colecao'])
                        aposta.append(row['Multiplo aposta'])
                        bloqueado.append(row['Bloqueado'])
                        tamanho.append(df_filtrado['GRADE'].iloc[0])
                        lista_estoque.append(estoque_ideal_programado)
                        reposicao.append(estoque_ideal_programado if (estoque_ideal_programado - estoque_filial) <= 0 else estoque_ideal_programado - estoque_filial)
                elif df_filtrado['Estoque_Ideal_Ajustado'].iloc[0] <= 0:
                    if row['Multiplo aposta'] > 0:
                        estoque_filial = df_filtrado['estoque_filial'].iloc[0]
                        estoque_ideal_programado = row[f'Estoque Mínimo {chave}']
                        lojas.append(row['filial'])
                        produto.append(row['PRODUTO'])
                        desc_produto.append(row['DESC_PRODUTO'])
                        cor_produto.append(row['cor_produto'])
                        desc_cor_produto.append(df_filtrado['desc_cor_produto'].iloc[0])
                        colecao.append(row['colecao'])
                        aposta.append(row['Multiplo aposta'])
                        bloqueado.append(row['Bloqueado'])
                        tamanho.append(df_filtrado['GRADE'].iloc[0])
                        lista_estoque.append(estoque_ideal_programado)
                        reposicao.append(estoque_ideal_programado if (estoque_ideal_programado - estoque_filial) <= 0 else estoque_ideal_programado - estoque_filial)
                    else:
                        estoque_filial = df_filtrado['estoque_filial'].iloc[0]
                        estoque_ideal_programado = row[f'Estoque Mínimo {chave}']
                        lojas.append(row['filial'])
                        produto.append(row['PRODUTO'])
                        desc_produto.append(row['DESC_PRODUTO'])
                        cor_produto.append(row['cor_produto'])
                        desc_cor_produto.append(df_filtrado['desc_cor_produto'].iloc[0])
                        colecao.append(row['colecao'])
                        aposta.append(row['Multiplo aposta'])
                        bloqueado.append(row['Bloqueado'])
                        tamanho.append(df_filtrado['GRADE'].iloc[0])
                        lista_estoque.append(estoque_ideal_programado)
                        reposicao.append(estoque_ideal_programado if (estoque_ideal_programado - estoque_filial) <= 0 else estoque_ideal_programado - estoque_filial)
                else:
                    if row['Multiplo aposta'] > 0:
                        estoque_filial = df_filtrado['estoque_filial'].iloc[0]
                        estoque_ideal_programado = round(row[f'Estoque Máximo {chave}'] * row['Multiplo aposta'])
                        lojas.append(row['filial'])
                        produto.append(row['PRODUTO'])
                        desc_produto.append(row['DESC_PRODUTO'])
                        cor_produto.append(row['cor_produto'])
                        desc_cor_produto.append(df_filtrado['desc_cor_produto'].iloc[0])
                        colecao.append(row['colecao'])
                        aposta.append(row['Multiplo aposta'])
                        bloqueado.append(row['Bloqueado'])
                        tamanho.append(df_filtrado['GRADE'].iloc[0])
                        lista_estoque.append(estoque_ideal_programado)
                        reposicao.append(estoque_ideal_programado if (estoque_ideal_programado - estoque_filial) <= 0 else estoque_ideal_programado - estoque_filial)
                    else:
                        estoque_filial = df_filtrado['estoque_filial'].iloc[0]
                        estoque_ideal_programado = row[f'Estoque Máximo {chave}']
                        lojas.append(row['filial'])
                        produto.append(row['PRODUTO'])
                        desc_produto.append(row['DESC_PRODUTO'])
                        cor_produto.append(row['cor_produto'])
                        desc_cor_produto.append(df_filtrado['desc_cor_produto'].iloc[0])
                        colecao.append(row['colecao'])
                        aposta.append(row['Multiplo aposta'])
                        bloqueado.append(row['Bloqueado'])
                        tamanho.append(df_filtrado['GRADE'].iloc[0])
                        lista_estoque.append(estoque_ideal_programado)
                        reposicao.append(estoque_ideal_programado if (estoque_ideal_programado - estoque_filial) <= 0 else estoque_ideal_programado - estoque_filial)

    elif row['Bloqueado'] == 1:
        pass
    else:
        print(f'Produto ({row["PRODUTO"]}) não cadastrado na planilha.')

def processar_reposicao(loja, dataframe_loja, diretorio):
    dataframe_loja['filial'] = dataframe_loja['filial'].astype('string')
    dataframe_loja['PRODUTO'] = dataframe_loja['PRODUTO'].astype('string')
    dataframe_loja['cor_produto'] = dataframe_loja['cor_produto'].astype('string')

    dataframe_loja.apply(Reposicao, axis=1)
    dict_reposicao = {'filial': lojas,
                    'PRODUTO': produto,
                    'DESC_PRODUTO': desc_produto,
                    'cor_produto': cor_produto,
                    'DESC_COR_PRODUTO': desc_cor_produto,
                    'Coleção': colecao,
                    'Multiplo aposta': aposta,
                    'Bloqueado': bloqueado,
                    'Grade': tamanho,
                    'Estoque Ideal - Programado': lista_estoque,
                    'Reposicao': reposicao}
    df_reposicao = pd.DataFrame(dict_reposicao)
    df_reposicao.to_excel(f'{diretorio}/OneDrive - Austral/Documentos - Millenium/Reposição/Reposição - {loja}.xlsx', index=False)
    print(f'Reposição da(o) {loja} feita com sucesso')


df = read_SQL(servidor, dados, usuario, senha, query)

filiais = ['LOJA1', 'LOJA2', 'LOJA3', 'LOJA4', 'LOJA5']

while True:
    print('Bem-Vindo ao sistema de reposição')
    print('Deseja fazer a reposição de qual loja? (Digite o número que corresponde a loja ou 6 para gerar de todas as lojas ao mesmo tempo)')
    print('1 - Loja 1')
    print('2 - Loja 2')
    print('3 - Loja 3')
    print('4 - Loja 4')
    print('5 - Loja 5')
    print('6 - Todas as lojas')
    print('7 - Digite esse caso queira sair.')

    resposta = int(input('Digite sua Resposta: '))

    if resposta == 1:
        lojas = []
        produto = []
        desc_produto = []
        cor_produto = []
        desc_cor_produto = []
        colecao = []
        aposta = []
        bloqueado = []
        tamanho = []
        lista_estoque = []
        reposicao = []
        filial = filiais[0]
        df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

        processar_reposicao(filial, df_loja, user_diretorio)

        print('Deseja fazer a reposição de outra loja ou deseja sair?? (Caso queira sair aperte 1, caso não, aperte qualquer outro número.)')

        saida = int(input('Por favor digite sua respota: '))

        if saida == 1:
            print('Obrigado e volte sempre...')
            time.sleep(3)
            break
        else:
            pass

    elif resposta == 2:
        lojas = []
        produto = []
        desc_produto = []
        cor_produto = []
        desc_cor_produto = []
        colecao = []
        aposta = []
        bloqueado = []
        tamanho = []
        lista_estoque = []
        reposicao = []
        filial = filiais[1]
        df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

        processar_reposicao(filial, df_loja, user_diretorio)

        print('Deseja fazer a reposição de outra loja ou deseja sair?? (Caso queira sair aperte 1, caso não, aperte qualquer outro número.)')

        saida = int(input('Por favor digite sua respota: '))

        if saida == 1:
            print('Obrigado e volte sempre...')
            time.sleep(3)
            break
        else:
            pass
    elif resposta == 3:
        lojas = []
        produto = []
        desc_produto = []
        cor_produto = []
        desc_cor_produto = []
        colecao = []
        aposta = []
        bloqueado = []
        tamanho = []
        lista_estoque = []
        reposicao = []
        filial = filiais[2]
        df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

        processar_reposicao(filial, df_loja, user_diretorio)

        print('Deseja fazer a reposição de outra loja ou deseja sair?? (Caso queira sair aperte 1, caso não, aperte qualquer outro número.)')

        saida = int(input('Por favor digite sua respota: '))

        if saida == 1:
            print('Obrigado e volte sempre...')
            time.sleep(3)
            break
        else:
            pass
    elif resposta == 4:
        lojas = []
        produto = []
        desc_produto = []
        cor_produto = []
        desc_cor_produto = []
        colecao = []
        aposta = []
        bloqueado = []
        tamanho = []
        lista_estoque = []
        reposicao = []
        filial = filiais[3]
        df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

        processar_reposicao(filial, df_loja, user_diretorio)

        print('Deseja fazer a reposição de outra loja ou deseja sair?? (Caso queira sair aperte 1, caso não, aperte qualquer outro número.)')

        saida = int(input('Por favor digite sua respota: '))

        if saida == 1:
            print('Obrigado e volte sempre...')
            time.sleep(3)
            break
        else:
            pass
    elif resposta == 5:
        lojas = []
        produto = []
        desc_produto = []
        cor_produto = []
        desc_cor_produto = []
        colecao = []
        aposta = []
        bloqueado = []
        tamanho = []
        lista_estoque = []
        reposicao = []
        filial = filiais[4]
        df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

        processar_reposicao(filial, df_loja, user_diretorio)

        print('Deseja fazer a reposição de outra loja ou deseja sair?? (Caso queira sair aperte 1, caso não, aperte qualquer outro número.)')

        saida = int(input('Por favor digite sua respota: '))

        if saida == 1:
            print('Obrigado e volte sempre...')
            time.sleep(3)
            break
        else:
            pass
    elif resposta == 6:
        for filial in filiais:
            if filial == 'LOJA PATIO HIGIENOPOLIS':
                lojas = []
                produto = []
                desc_produto = []
                cor_produto = []
                desc_cor_produto = []
                colecao = []
                aposta = []
                bloqueado = []
                tamanho = []
                lista_estoque = []
                reposicao = []
                df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

                processar_reposicao(filial, df_loja, user_diretorio)
            elif filial == 'LOJA IGUATEMI SP':
                lojas = []
                produto = []
                desc_produto = []
                cor_produto = []
                desc_cor_produto = []
                colecao = []
                aposta = []
                bloqueado = []
                tamanho = []
                lista_estoque = []
                reposicao = []
                df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

                processar_reposicao(filial, df_loja, user_diretorio)
            elif filial == 'IGUATEMI ALPHAVILLE':
                lojas = []
                produto = []
                desc_produto = []
                cor_produto = []
                desc_cor_produto = []
                colecao = []
                aposta = []
                bloqueado = []
                tamanho = []
                lista_estoque = []
                reposicao = []
                df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

                processar_reposicao(filial, df_loja, user_diretorio)
            elif filial == 'LOJA IGUATEMI JK':
                lojas = []
                produto = []
                desc_produto = []
                cor_produto = []
                desc_cor_produto = []
                colecao = []
                aposta = []
                bloqueado = []
                tamanho = []
                lista_estoque = []
                reposicao = []
                df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

                processar_reposicao(filial, df_loja, user_diretorio)
            elif filial == 'LOJA MORUMBI':
                lojas = []
                produto = []
                desc_produto = []
                cor_produto = []
                desc_cor_produto = []
                colecao = []
                aposta = []
                bloqueado = []
                tamanho = []
                lista_estoque = []
                reposicao = []
                df_loja = pd.read_excel(f'{user_diretorio}/caminho até o estoque da loja para leitura e análise/{filial} - Variaveis Reposicao.xlsx')

                processar_reposicao(filial, df_loja, user_diretorio)
        print('Reposição de todas as lojas criadas com sucesso.')
        print(f'Acesse o caminho "{user_diretorio}/caminho para onde as planilhas de reposição das lojas estão guardados/Reposição" para verificar as planilhas')
        print('Obrigado e volte sempre...')
        time.sleep(3)
        break

    elif resposta == 7:
        print('Obrigado e volte sempre...')
        time.sleep(3)
        break
    else:
        print(f'O número: {resposta} não existe.')
        print('Por favor digite um número válido!!')