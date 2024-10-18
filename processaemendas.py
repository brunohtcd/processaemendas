import xlwings as xw
import pandas as pd
from dataclasses import dataclass
from typing import Optional
from enum import Enum


# Constantes para nomes de planilhas
PLANILHA_RELATOR = 'Relator - Coletivas Apropriação'
PLANILHA_CRITERIO = 'Critério IU 6'
PLANILHA_BANCADA = 'Bancada Impositiva - Consulta'
PLANILHA_BOTAO = 'Sheet1'

class Parecer(Enum):
    """
    Enum que representa os diferentes pareceres que podem ser atribuídos a uma emenda.
    """
    APROVACAO = 1
    APROVACAO_PARCIAL = 13 
    REJEICAO = 38 


class Decisao(Enum):
    """
    Enum que representa as diferentes decisões que podem ser tomadas para uma emenda.
    """
    APROVACAO = 1
    APROVACAO_PARCIAL = 2
    REJEICAO = 3

# Dicionário que mapeia cada parecer para a decisão correspondente
DECISAO_PARA_PARECER = {
    Decisao.APROVACAO: Parecer.APROVACAO,
    Decisao.APROVACAO_PARCIAL: Parecer.APROVACAO_PARCIAL,
    Decisao.REJEICAO: Parecer.REJEICAO
}


# Função para verificar se a emenda está na 'Bancada Impositiva - Consulta'
def checa_bancada_impositiva(emenda: str, df_bancada: pd.DataFrame) -> bool:
    """
    Verifica se uma emenda está presente na planilha 'Bancada Impositiva - Consulta'.
    
    Args:
        emenda (str): O código da emenda a ser verificada.
        df_bancada (pd.DataFrame): DataFrame contendo as informações da bancada impositiva.
    
    Returns:
        bool: True se a emenda estiver presente, False caso contrário.
    """
    return emenda in df_bancada['Emenda'].values



# Função para calcular a 'Decisão' de acordo com a lógica fornecida
def calcula_decisao(linha: pd.Series, df_bancada: pd.DataFrame) -> Decisao:
    """
    Calcula a decisão de uma emenda com base em regras de atendimento setorial e presença na bancada impositiva.
    Em resumo, a fórmula retornará:
    1 se a soma em 'Atendimento Setorial' for igual à soma em 'Atendimento Setorial' (e ambas não forem zero), ou seja, se valor solicitado for = aprovado, então emenda aprovada;
    2 se a soma em 'Atendimento Setorial' for diferente de 'Atendimento Setorial' (mas não zero), ou se a emenda for encontrado na planilha de impositivas 
      (ou seja, aprovado parcial se atendido for menor que solicitado ou, se valor atendido discricionária for zero, se a emenda tiver parcela impositiva)
    3 se a soma em 'Atendimento Setorial' for zero e A2 não for encontrado na outra planilha (ou seja, emenda rejeitada)
    
    Args:
        linha (pd.Series): Objeto Emend contendo a linha com os dados da emenda.
        df_bancada (pd.DataFrame): DataFrame contendo as informações da bancada impositiva.
    
    Returns:
        Parecer: Enum indicando o parecer (APROVACAO, APROVACAO_PARCIAL, REJEICAO).
    """
    if linha['Soma Atendimento Setorial'] != 0:
        if linha['Soma Atendimento Setorial'] == linha['Soma Valor Solicitado']:
            return Decisao.APROVACAO
        else:
            return Decisao.APROVACAO_PARCIAL
    else:
        if checa_bancada_impositiva(linha['Emenda'], df_bancada):
            return Decisao.APROVACAO_PARCIAL
        else:
            return Decisao.REJEICAO


def calcula_id_uso(linha: pd.Series, df_criterio: pd.DataFrame) -> int:
    """
    Calcula o ID de uso para uma emenda com base em critérios de UO e funcional presentes na tabela de critério.
    
    Args:
        linha (pd.Series): Objeto Emend contendo a linha com os dados da emenda.
        df_criterio (pd.DataFrame): DataFrame contendo os critérios de uso.
    
    Returns:
        int: ID de uso (6 ou 0).
    """

    count_uo = df_criterio['UO'].eq(int(linha['UO'])).sum()
    acao = linha['Funcional'][-4:]
    count_acao = df_criterio['Ação'].eq(acao).sum()
    
    # Condicional para retornar 6 ou 0
    if count_uo == 1 and count_uo == 1:
        return 6
    else:
        return 0

def calcula_parecer(decisao_parecer: Decisao) -> Parecer:
    """
    Mapeia a decisão para um parecer correspondente.
    
    Args:
        decisao_parecer (Decisao): Enum indicando a Decisao.
    
    Returns:
        Decisao: Enum indicando o parecer correspondente a Decisao.
    """
    return DECISAO_PARA_PARECER[decisao_parecer]

def formata_dataframe(df):
    """
      Atualiza o dataframe original, mantendo apenas até a última linha com valor numérico.
      Remove a última linha de contabilização e possíveis linhas vazias ao final da tabela
      Preenche células mescladas (NaN) com o último valor válido
    Args:
        df (pd.DataFrame): DataFrame a ser formatado.
    
    Returns:
        pd.DataFrame: DataFrame formatado.
    """
    
    ultima_linha_numerica = df[df['Emenda'].apply(lambda x: pd.to_numeric(x, errors='coerce')).notna()].index[-1]
    df = df.loc[:ultima_linha_numerica].copy()
    
    df = df.ffill()
    return df

def somas_por_emendas(df_relator):
    """
    Calcula as somas de atendimento setorial e valor solicitado para cada emenda.
    
    Args:
        df_relator (pd.DataFrame): DataFrame contendo os dados das emendas.
    
    Returns:
        pd.DataFrame: DataFrame atualizado com colunas de soma para atendimento setorial e valor solicitado.
    """
    
    somas_atendimento_setorial = df_relator.groupby('Emenda')['Atendimento Setorial'].sum()
    somas_valor_solicitado = df_relator.groupby('Emenda')['Valor Solicitado'].sum()

    # Adiciona essas somas como novas colunas ao dataframe original, usando o valor da emenda como referência
    df_relator = df_relator.merge(somas_atendimento_setorial.rename('Soma Atendimento Setorial'), 
                              left_on='Emenda', right_index=True, how='left')
    df_relator = df_relator.merge(somas_valor_solicitado.rename('Soma Valor Solicitado'), 
                              left_on='Emenda', right_index=True, how='left')
    return df_relator

def somas_parcelas_impositivas(df_bancada):
    """
    Calcula a soma dos valores solicitados para cada emenda na bancada impositiva.
    
    Args:
        df_bancada (pd.DataFrame): DataFrame contendo os dados da bancada impositiva.
    
    Returns:
        pd.DataFrame: DataFrame atualizado com a coluna de soma de valores solicitados.
    """
    somas_parcela_impositiva = df_bancada.groupby('Emenda')['Valor Solicitado'].sum()
    df_bancada = df_bancada.merge(somas_parcela_impositiva.rename('Soma Valor Solicitado'), 
                              left_on='Emenda', right_index=True, how='left')
    
    return df_bancada

def mapeia_parcelas_impositivas(df_bancada,df_relator):
    """
    Mapeia parcelas impositivas do DataFrame da bancada para o DataFrame do relator.
    
    Args:
        df_bancada (pd.DataFrame): DataFrame contendo os dados da bancada impositiva.
        df_relator (pd.DataFrame): DataFrame contendo os dados do relator.
    
    Returns:
        pd.DataFrame: DataFrame do relator atualizado com a coluna de parcelas impositivas.
    """
    # Criar um dicionário de mapeamento a partir do DataFrame df_bancada
    mapeamento = dict(zip(df_bancada['Emenda'], df_bancada['Soma Valor Solicitado']))

    # Atualizar a coluna 'Tem parcela impositiva?' do df_relator usando o método .map()
    # Caso não haja correspondência, preenche com '-'
    df_relator['Tem parcela impositiva?'] = df_relator['Emenda'].map(mapeamento).fillna('-')

    return df_relator

def processa_emenda(linha: pd.Series, df_criterio: pd.DataFrame, df_bancada: pd.DataFrame, emenda_bancada: bool = False) -> pd.Series:
    """
    Processa uma linha do DataFrame, calculando ID de uso, decisão do parecer e parecer padrão.
    
    Args:
        row (pd.Series): Linha do DataFrame contendo os dados da emenda.
        df_criterio (pd.DataFrame): DataFrame contendo os critérios de uso.
        df_bancada (pd.DataFrame): DataFrame contendo os dados da bancada impositiva.
    
    Returns:
        pd.Series: Linha atualizada com as novas colunas calculadas.
    """
   
    linha['ID Uso'] = calcula_id_uso(linha, df_criterio)
    
    if emenda_bancada:
        decisao = Decisao.APROVACAO
    else:
        decisao = calcula_decisao(linha, df_bancada)
    
    linha['Decisão Parecer'] = decisao.value
    parecer = calcula_parecer(decisao)
    linha['Parecer Padrão'] = parecer.value
    linha['Valor'] = linha['Atendimento Setorial']
    
    return linha


def processa_emendas(df: pd.DataFrame, df_criterio: pd.DataFrame, df_bancada: pd.DataFrame, funcao_processamento: callable, emenda_bancada: bool = False) -> pd.DataFrame:
    """
    Processa todas as emendas no DataFrame do relator.
    
    Args:
        df_relator (pd.DataFrame): DataFrame contendo os dados das emendas do relator.
        df_criterio (pd.DataFrame): DataFrame contendo os critérios de uso.
        df_bancada (pd.DataFrame): DataFrame contendo os dados da bancada impositiva.
    
    Returns:
        pd.DataFrame: DataFrame do relator atualizado com as novas colunas calculadas.
    """
    return df.apply(funcao_processamento, axis=1, args=(df_criterio, df_bancada, emenda_bancada))

def main():
    """
    Função principal para carregar os dados das planilhas, processar as emendas e atualizar as planilhas.
    """
    # Conectar ao livro de trabalho ativo no xlwings
    wb = xw.Book.caller()  # Isso conecta ao arquivo Excel aberto pelo xlwings
    sheet_plansheet = wb.sheets[PLANILHA_BOTAO]
    sheet_plansheet.range('A1').value = 'Gerando planilha Lexor...'

    # Carregar as abas 
    sheet_apropriacao = wb.sheets[PLANILHA_RELATOR]
    sheet_criterio = wb.sheets[PLANILHA_CRITERIO]
    sheet_bancada = wb.sheets[PLANILHA_BANCADA]
    # Carregar as abas como DataFrames pandas
    df_relator =  sheet_apropriacao.used_range.options(pd.DataFrame, index=False, header=True).value
    df_criterio = sheet_criterio.used_range.options(pd.DataFrame, index=False, header=True).value
    df_bancada = sheet_bancada.used_range.options(pd.DataFrame, index=False, header=True).value

    # Formate dataframes
    df_bancada = formata_dataframe(df_bancada)
    df_relator = formata_dataframe(df_relator)
    # Realiza calulos auxiliares
    df_bancada = somas_parcelas_impositivas(df_bancada)
    df_relator = somas_por_emendas(df_relator)
    df_relator = mapeia_parcelas_impositivas(df_bancada,df_relator)
    # Processa todas as emendas 
    df_relator = processa_emendas(df=df_relator, df_criterio=df_criterio, df_bancada=df_bancada, funcao_processamento=processa_emenda,emenda_bancada=False)
    # processa emendas de banada
    df_bancada = processa_emendas(df=df_bancada, df_criterio=df_criterio, df_bancada=None, funcao_processamento=processa_emenda, emenda_bancada=True)
       
    # Escrever o DataFrame `df_relator de volta na planilha original
    sheet_apropriacao.range('A1').options(index=False).value = df_relator.iloc[:, :-2] # 'reset_index' evita o índice do Dataframe. Descartando as duas últimas colunas de soma criadas anteriormente
    # Escrever o DataFrame `df_relator de volta na planilha original
    sheet_bancada.range('A1').options(index=False).value = df_bancada.iloc[:,:-1] #Descartando a última coluna de soma de valor solicitado por emenda criada anteriormente

    sheet_plansheet.range('A1').value = 'Planilha Lexor gerada com sucesso!'



if __name__ == "__main__":
    xw.Book("processaemendas.xlsm").set_mock_caller()
    main()
