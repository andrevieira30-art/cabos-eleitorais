import os
import oracledb
from dotenv import load_dotenv

load_dotenv()

ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")


def conectar_oracle():
    try:

        dsn = """
        (description=
            (retry_count=20)
            (retry_delay=3)
            (address=
                (protocol=tcps)
                (port=1522)
                (host=adb.sa-saopaulo-1.oraclecloud.com)
            )
            (connect_data=
                (service_name=g5087928fba57e8_orclapi_low.adb.oraclecloud.com)
            )
            (security=
                (ssl_server_dn_match=yes)
            )
        )
        """

        conexao = oracledb.connect(
            user=ORACLE_USER,
            password=ORACLE_PASSWORD,
            dsn=dsn
        )

        print("Conectado ao Oracle com sucesso!")
        return conexao

    except Exception as erro:
        print("Erro ao conectar no Oracle:", erro)
        return None