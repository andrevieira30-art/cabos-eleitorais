import os
import oracledb

PASTA_WALLET = r"C:\Users\cliente\AppData\Local\Programs\Python\Python312\portal_documentos_oracle\CabosEleitorais"

print("tnsnames existe?", os.path.exists(os.path.join(PASTA_WALLET, "tnsnames.ora")))
print("ewallet.pem existe?", os.path.exists(os.path.join(PASTA_WALLET, "ewallet.pem")))
print("sqlnet existe?", os.path.exists(os.path.join(PASTA_WALLET, "sqlnet.ora")))

try:
    print("Tentando conexão...")
    conn = oracledb.connect(
        user="ADMIN",
        password="Aduser@#902060",
        dsn="orclapi_high",
        config_dir=PASTA_WALLET,
        wallet_location=PASTA_WALLET,
        wallet_password="Aduser902060",
        tcp_connect_timeout=10
    )
    print("Conectou com sucesso!")
    cur = conn.cursor()
    cur.execute("select sysdate from dual")
    print(cur.fetchone())
    cur.close()
    conn.close()
except Exception as e:
    print("Erro:", e)