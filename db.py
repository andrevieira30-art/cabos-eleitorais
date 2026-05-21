import os
import oracledb
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")


def conectar_oracle():
    try:

        wallet_path = os.path.join(BASE_DIR, "wallet")

        conexao = oracledb.connect(
            user=ORACLE_USER,
            password=ORACLE_PASSWORD,
            dsn="orclapi_low",
            config_dir=wallet_path,
            wallet_location=wallet_path,
            wallet_password=ORACLE_PASSWORD
        )

        print("Conectado ao Oracle com sucesso!")
        return conexao

    except Exception as erro:
        print("Erro ao conectar no Oracle:", erro)
        return None