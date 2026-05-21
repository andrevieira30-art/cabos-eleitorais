import os
import oracledb
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")
ORACLE_DSN = os.getenv("ORACLE_DSN")


def conectar_oracle():
    try:
        conexao = oracledb.connect(
            user=ORACLE_USER,
            password=ORACLE_PASSWORD,
            dsn=ORACLE_DSN
        )

        print("Conectado no Oracle com sucesso")
        return conexao

    except oracledb.Error as erro:
        print("Erro ao conectar no Oracle:", erro)
        return None