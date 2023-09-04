from sqlalchemy import create_engine
from configparser import ConfigParser
import pyodbc as sqldriver
from urllib.parse import quote
import pandas as pd


config = ConfigParser()
config.read('conf.ini')


def get_conn(machineName, name=None):
    server = config[machineName]['sql_server']
    port = config[machineName]['sql_port']
    table = config[machineName]['sql_table']
    username = config[machineName]['sql_username']
    password = quote(config[machineName]['sql_password'])
    sqlalchemy_driver = config[machineName]['sql_sqlalchemy_driver']

    conn = [
        'mssql+pymssql://username:password@server:port/table',
        'mssql+pyodbc://username:password@server:port/table?driver=sqlalchemy_driver'
    ]

    conn = (conn[0]
            .replace('username', username)
            .replace('server', server)
            .replace('port', port)
            .replace('table', table)
            .replace('username', username)
            .replace('password', password)
            .replace('sqlalchemy_driver', sqlalchemy_driver)
            )
    engine = create_engine(conn)
    return engine


conn = get_conn('mssql_chemtest')
df = pd.read_sql('select top(10) * FROM [ChemiPrimary].[dbo].[AttnChemicalList]', con=conn)