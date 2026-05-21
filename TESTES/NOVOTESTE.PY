import os
import oracledb

PASTA_WALLET = r"C:\Users\cliente\AppData\Local\Programs\Python\Python312\portal_documentos_oracle\CabosEleitorais"

print("tnsnames existe?", os.path.exists(os.path.join(PASTA_WALLET, "tnsnames.ora")))
print("ewallet.pem existe?", os.path.exists(os.path.join(PASTA_WALLET, "ewallet.pem")))
print("sqlnet existe?", os.path.exists(os.path.join(PASTA_WALLET, "sqlnet.ora")))

aliases = [
    "ORCLAPI_HIGH",
    "ORCLAPI_MEDIUM",
    "ORCLAPI_LOW",
    "ORCLAPI_TP",
    "ORCLAPI_TPURGENT"
]

for alias in aliases:
    try:
        print(f"\nTestando alias: {alias}")

        conn = oracledb.connect(
            user="ADMIN",
            password="Aduser@#902060",
            dsn=alias,
            config_dir=PASTA_WALLET,
            wallet_location=PASTA_WALLET,
            wallet_password="Aduser902060",
            tcp_connect_timeout=10,
            retry_count=0
        )

        print("Conectou com sucesso!")
        cur = conn.cursor()
        cur.execute("SELECT SYSDATE FROM DUAL")
        print("Resultado:", cur.fetchone())

        cur.close()
        conn.close()
        break

    except Exception as e:
        print("Falhou:", repr(e))