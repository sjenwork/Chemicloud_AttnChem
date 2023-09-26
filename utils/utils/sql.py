from sqlalchemy import create_engine
from sqlalchemy import create_engine, Integer, DateTime, NVARCHAR, Float
from configparser import ConfigParser
from urllib.parse import quote_plus
import pathlib
import socket

def read_conf():
    config = ConfigParser()
    configfile = pathlib.Path(__file__).parents[1]
    config.read(configfile/'config/conf.ini')
    return config

def connSQL(db='chemiBD_Test'):
    conf = read_conf()[db]
    # uri = 'mssql+pymssql://username:password@server:port/db'
    hostname = socket.gethostname()
    uri = (
        'mssql+pymssql://username:password@server:port/db?charset=utf8'
        if hostname != 'jenMBP14.local'
        else 'mssql+pyodbc://username:password@server:port/db?driver=sqlalchemy_driver;charset=utf8'
    )        
    uri = (
        uri
        .replace('username', conf['sql_username'])
        .replace('password', quote_plus(conf['sql_password']))
        .replace('server', conf['sql_server'])
        .replace('port', conf['sql_port'])
        .replace('db', conf['sql_dbname'])
    )
    engine = create_engine(uri)
    return engine

def createSchema(dfparam):
    dtypedict = {}
    for i,j in zip(dfparam.columns, dfparam.dtypes):
        if "object" in str(j).lower():
            dtypedict.update({i: NVARCHAR()})

        if "string" in str(j).lower():
            dtypedict.update({i: NVARCHAR()})

        if "datetime" in str(j).lower():
            dtypedict.update({i: DateTime()})

        if "float" in str(j).lower():
            dtypedict.update({i: Float(precision=6, asdecimal=True)})

        if "int" in str(j).lower():
            dtypedict.update({i: Integer()})

    return dtypedict
