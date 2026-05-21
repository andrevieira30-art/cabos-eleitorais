import oracledb

params = oracledb.ConnectParams(config_dir=r"C:\Users\cliente\AppData\Local\Programs\Python\Python312\portal_documentos_oracle\CabosEleitorais")
print(params.get_network_service_names())