import os
import oracledb

PASTA_WALLET = r"C:\ORCLAPI_wallet_nova"

print("tnsnames existe?", os.path.exists(os.path.join(PASTA_WALLET, "tnsnames.ora")))
print("ewallet.pem existe?", os.path.exists(os.path.join(PASTA_WALLET, "ewallet.pem")))

params = oracledb.ConnectParams(config_dir=PASTA_WALLET)
print("Aliases:", params.get_network_service_names())

try:
    conn = oracledb.connect(
        user="ADMIN",
        password="Aduser@#902060",
        dsn="ORCLAPI_HIGH",
        config_dir=PASTA_WALLET,
        wallet_location=PASTA_WALLET,
        wallet_password="Aduser605090",
        tcp_connect_timeout=10,
        retry_count=0
    )
    print("Conectou com sucesso!")
    cur = conn.cursor()
    cur.execute("select sysdate from dual")
    print("Resultado:", cur.fetchone())
    cur.close()
    conn.close()
except Exception as e:
    print("Erro:", repr(e))