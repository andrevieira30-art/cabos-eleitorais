import os
import oracledb
from dotenv import load_dotenv

load_dotenv()

ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")


def conectar_oracle():
    try:

        dsn = "adb.sa-saopaulo-1.oraclecloud.com:1522/g5087928fba57e8_orclapi_low.adb.oraclecloud.com"

        conexao = oracledb.connect(
            user=ORACLE_USER,
            password=ORACLE_PASSWORD,
            dsn=dsn,
            protocol="tcps"
        )

        print("Conectado ao Oracle com sucesso!")
        return conexao

    except Exception as erro:
        print("Erro ao conectar no Oracle:", erro)
        return None