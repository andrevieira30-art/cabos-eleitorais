import os
import oracledb
from dotenv import load_dotenv

load_dotenv()

ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")
ORACLE_DSN = os.getenv("ORACLE_DSN")


def conectar_oracle():
    try:
        conexao = oracledb.connect(
            user=ORACLE_USER,
            password=ORACLE_PASSWORD,
            host="adb.sa-saopaulo-1.oraclecloud.com",
            port=1522,
            service_name=ORACLE_DSN,
            protocol="tcps"
        )

        print("Conectado ao Oracle com sucesso!")
        return conexao

    except Exception as erro:
        print("Erro ao conectar no Oracle:", erro)
        return None