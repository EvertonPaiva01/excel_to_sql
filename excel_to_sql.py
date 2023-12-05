#Importando bibliotecas
import pandas as pd
from datetime import date
from sqlalchemy import create_engine
import pymysql

#Iniciando protocolo de comunicação com o MySQL
pymysql.install_as_MySQLdb()

#Carregando arquivos do excel
arquivo = '_BaseContasAnalisadas.xlsx'
pasta = 'C:/Users/Everton Paiva/Documents/Neoenergia/_BaseContasAnalisadas/'

#Criando o DataFrame
df = pd.read_excel(pasta+arquivo)

#Função para se conectar ao banco de dados
def mysql_connection(host, user, passwd, database=None):
    engine = create_engine(f'mysql+pymysql://{user}:{passwd}@{host}/{database}')
    return engine.connect()

#Conexão como Banco de Dados
#Meu nome não é Johnny
ip = '192.168.87.95'
user = 'EMLURB'
password = 'gipemlurb123'
bank = 'db_tois_analisados'

connection = mysql_connection(ip, user, password, bank)

#Filtrando a tabela e selecionando apenas as colunas que serão utilizadas no banco de dados
tabela_filtrada = df[['CONTA CONTRATO DA FATURA', 'TOI', 'INSTALAÇÃO', ' MONTANTE', 'STATUS DA ANALISE(ATESTADO)',
       'Arquivo', 'Data Analise']].copy()

#Atribuindo todos os filtros para a tabela final
tabela = tabela_filtrada.copy()
#Criando uma cópia da tabela com todos os filtros para ser adicionada ao banco de dados
df1 = tabela.copy()

#Redefinindo o nome das colunas
df1.columns = ['ContaContratoFatura', 'TOI', 'Instalacao', 'Montante', 'StatusAnalise',
       'Arquivo', 'DataAnalise']

#Tranforando NaN para vazio e inserindo no banco de dados
df1.fillna('').to_sql('tois',connection,if_exists='append',index=False)