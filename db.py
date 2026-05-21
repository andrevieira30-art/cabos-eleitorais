import os
import oracledb
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")
ORACLE_DSN = os.getenv("ORACLE_DSN")
ORACLE_WALLET_DIR = os.getenv("ORACLE_WALLET_DIR")
ORACLE_WALLET_PASSWORD = os.getenv("ORACLE_WALLET_PASSWORD")


def conectar_oracle():
    try:
        if not ORACLE_WALLET_DIR:
            print("Erro: ORACLE_WALLET_DIR não foi carregado do .env")
            return None

        conexao = oracledb.connect(
            user=ORACLE_USER,
            password=ORACLE_PASSWORD,
            dsn=ORACLE_DSN,
            config_dir=ORACLE_WALLET_DIR,
            wallet_location=ORACLE_WALLET_DIR,
            wallet_password=ORACLE_WALLET_PASSWORD,
            tcp_connect_timeout=10,
            retry_count=0
        )
        return conexao

    except oracledb.Error as erro:
        print("Erro ao conectar no Oracle:", erro)
        return None