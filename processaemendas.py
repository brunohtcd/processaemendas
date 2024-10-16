import xlwings as xw
import pandas as pd



df = ""
df_criterio = ""
df_bancada = ""



# Função para verificar se a emenda está na 'Bancada Impositiva - Consulta'
def checa_bancada_impositiva(emenda, df_bancada):
    
    # Verifica se a emenda está presente na coluna 'Emenda' da aba 'Bancada Impositiva - Consulta'
    return emenda in df_bancada['Emenda'].values


# Função para calcular a 'Decisão' de acordo com a lógica fornecida
def calcula_decisao(row, df_bancada):
    # Se a soma de 'Atendimento Setorial' não for zero
    if row['Soma Atendimento Setorial'] != 0:
        # Se a soma de 'Atendimento Setorial' for igual à soma de 'Valor Solicitado'
        if row['Soma Atendimento Setorial'] == row['Soma Valor Solicitado']:
            return 1  # Retorna 1 se forem iguais
        else:
            return 2  # Retorna 2 se forem diferentes
    else:
        # Se a soma de 'Atendimento Setorial' for zero, verificar a presença da emenda na aba 'Bancada Impositiva - Consulta'
        if checa_bancada_impositiva(row['Sequencial da Emenda'], df_bancada):
            return 2  # Retorna 2 se a emenda for encontrada na 'Bancada Impositiva - Consulta'
        else:
            return 3  # Retorna 3 se a emenda não for encontrada na 'Bancada Impositiva - Consulta'


# Função para simular a fórmula de 'ID Uso'
def calcula_id_uso(row,df_criterio):
        
    # Contagem condicional em 'Critério IU 6'
    UO = int(row['UO'])
    count_x = df_criterio['UO'].eq(UO).sum()  # Equivalente a CONT.SE('Critério IU 6'!A:A;X2)
    last_4_of_y = row['Funcional'][-4:]  # Equivalente a DIREITA(Y2;4)
    count_y = df_criterio['Ação'].eq(last_4_of_y).sum()  # Equivalente a CONT.SE('Critério IU 6'!B:B;DIREITA(Y2;4))
    
    # Condicional para retornar 6 ou 0
    if count_x == 1 and count_y == 1:
        return 6
    else:
        return 0

def calcula_parecer(row):
    # Simula a lógica da fórmula =SE(AD2=1;1;SE(AD2=2;13;38)) onde AD2 é a coluna 'Decisão'
    if row['Decisão Parecer'] == 1:
        return 1
    elif row['Decisão Parecer'] == 2:
        return 13
    else:
        return 38

def formata_dataframe(df_relator):
    # Atualizar o dataframe original, mantendo apenas até a última linha com valor numérico.
    # Remove a última linha de contabilização e possíveis linhas vazias ao final da tabela
    ultima_linha_numerica = df_relator[df_relator['Emenda'].apply(lambda x: pd.to_numeric(x, errors='coerce')).notna()].index[-1]
    df_relator = df_relator.loc[:ultima_linha_numerica].copy()
    # Preencher células mescladas (NaN) com o último valor válido
    df_relator = df_relator.ffill()
    return df_relator

def somas_por_emendas(df_relator):
    
    somas_atendimento_setorial = df_relator.groupby('Emenda')['Atendimento Setorial'].sum()
    somas_valor_solicitado = df_relator.groupby('Emenda')['Valor Solicitado'].sum()

    # Adiciona essas somas como novas colunas ao dataframe original, usando o valor da emenda como referência
    df_relator = df_relator.merge(somas_atendimento_setorial.rename('Soma Atendimento Setorial'), 
                              left_on='Emenda', right_index=True, how='left')
    df_relator = df_relator.merge(somas_valor_solicitado.rename('Soma Valor Solicitado'), 
                              left_on='Emenda', right_index=True, how='left')
    return df_relator

def somas_parcelas_impositivas(df_bancada):

    somas_parcela_impositiva = df_bancada.groupby('Emenda')['Valor Solicitado'].sum()
    df_bancada = df_bancada.merge(somas_parcela_impositiva.rename('Soma Valor Solicitado'), 
                              left_on='Emenda', right_index=True, how='left')
    
    return df_bancada

def mapeia_parcelas_impositivas(df_bancada,df_relator):
    
    # Criar um dicionário de mapeamento a partir do DataFrame df_bancada
    mapeamento = dict(zip(df_bancada['Emenda'], df_bancada['Soma Valor Solicitado']))

    # Atualizar a coluna 'Tem parcela impositiva?' do df_relator usando o método .map()
    # Caso não haja correspondência, preenche com '-'
    df_relator['Tem parcela impositiva?'] = df_relator['Emenda'].map(mapeamento).fillna('-')

    return df_relator
    

def main():
    # Conectar ao livro de trabalho ativo no xlwings
    wb = xw.Book.caller()  # Isso conecta ao arquivo Excel aberto pelo xlwings
    sheet_plansheet = wb.sheets['Sheet1']
    sheet_plansheet.range('A1').value = 'Gerando planilha Lexor...'
  

    # Carregar as abas como DataFrames
    sheet_apropriacao = wb.sheets['Relator - Coletivas Apropriação']
    sheet_criterio = wb.sheets['Critério IU 6']
    sheet_bancada = wb.sheets['Bancada Impositiva - Consulta']

    df_relator =  sheet_apropriacao.used_range.options(pd.DataFrame, index=False, header=True).value
    df_criterio = sheet_criterio.used_range.options(pd.DataFrame, index=False, header=True).value
    df_bancada = sheet_bancada.used_range.options(pd.DataFrame, index=False, header=True).value
    df_bancada = df_bancada.ffill()

    df_bancada = somas_parcelas_impositivas(df_bancada)
    df_relator = formata_dataframe(df_relator)
    df_relator = somas_por_emendas(df_relator)
    df_relator = mapeia_parcelas_impositivas(df_bancada,df_relator)

    # LEMBRAR DE SOLICITAR MUDANÇA DO NOME DA COLUNA Funcional!!!!! ######################
    
    # Aplicando a função para calcular 'ID Uso'
    df_relator['ID Uso'] = df_relator.apply(lambda row: calcula_id_uso(row,df_criterio), axis=1)
    # Aplicando a função para calcular 'Decisão Parecer'
    df_relator['Decisão Parecer'] = df_relator.apply(lambda row: calcula_decisao(row, df_bancada), axis=1)
    # Aplicando a função para calcular 'Parecer'
    df_relator['Parecer Padrão'] = df_relator.apply(lambda row: calcula_parecer(row), axis=1)
    
    
    # Escrever o DataFrame `df_lexor` de volta a partir da coluna "Sequencial da Emenda"
    sheet_apropriacao.range('A1').options(index=False).value = df_relator.iloc[:, :-2] # 'reset_index' evita o índice do Dataframe. Descartando as duas últimas colunas de soma criadas anteriormente

    sheet_plansheet.range('A1').value = 'Planilha Lexor gerada com sucesso!'





if __name__ == "__main__":
    xw.Book("processaemendas.xlsm").set_mock_caller()
    main()
