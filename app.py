import os
import secrets
import smtplib
import pytz
from werkzeug.security import generate_password_hash, check_password_hash
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO
from datetime import datetime
from functools import wraps


from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from dotenv import load_dotenv
import oracledb
from openpyxl import Workbook

from db import conectar_oracle

# =========================
# Configuração e helpers
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
MAIL_FROM = os.getenv("MAIL_FROM")
BASE_URL = os.getenv("BASE_URL", "http://127.0.0.1:5000")

ADMIN_USER = os.getenv("ADMIN_USER")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD")

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "chave_padrao")

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "minha_chave_secreta_123")

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("usuario_logado"):
            flash("Faça login para acessar o sistema.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

@app.template_filter("data_br")
def data_br(data):
    if not data:
        return "-"

    fuso_brasilia = pytz.timezone("America/Sao_Paulo")

    if data.tzinfo is None:
        data = pytz.utc.localize(data)

    data_brasilia = data.astimezone(fuso_brasilia)

    return data_brasilia.strftime("%d/%m/%Y %H:%M")

def perfil_required(*tipos_permitidos):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not session.get("usuario_logado"):
                flash("Faça login para acessar o sistema.", "warning")
                return redirect(url_for("login"))

            tipo_usuario = session.get("tipo_acesso")

            if tipo_usuario not in tipos_permitidos:
                flash("Você não tem permissão para acessar esta área.", "danger")
                return redirect(url_for("home"))

            return f(*args, **kwargs)
        return decorated_function
    return decorator


def converter_data(data_str):
    if not data_str:
        return None
    try:
        return datetime.strptime(data_str, "%Y-%m-%d").date()
    except ValueError:
        return None


def gerar_token_convite():
    return secrets.token_urlsafe(32)


def enviar_email_convite(destinatario, nome_cabo, link):
    assunto = "Convite para concluir seu cadastro"

    corpo_texto = f"""
Olá,

Você recebeu um convite para concluir seu cadastro.

Responsável pelo convite: {nome_cabo}

Acesse o link abaixo para continuar:
{link}

Se você não reconhece este convite, ignore esta mensagem.
"""

    corpo_html = f"""
    <html>
        <body>
            <h3>Olá!</h3>
            <p>Você recebeu um convite para concluir seu cadastro.</p>
            <p><strong>Responsável pelo convite:</strong> {nome_cabo}</p>
            <p>Acesse o link abaixo para continuar:</p>
            <p><a href="{link}">{link}</a></p>
            <p>Se você não reconhece este convite, ignore esta mensagem.</p>
        </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = assunto
    msg["From"] = MAIL_FROM
    msg["To"] = destinatario

    msg.attach(MIMEText(corpo_texto, "plain", "utf-8"))
    msg.attach(MIMEText(corpo_html, "html", "utf-8"))

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as servidor:
        servidor.starttls()
        servidor.login(SMTP_USER, SMTP_PASSWORD)

        recusados = servidor.sendmail(
            MAIL_FROM,
            [destinatario],
            msg.as_string()
        )

        print("RECUSADOS SMTP:", recusados)
        return recusados
    
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("usuario_logado"):
            flash("Faça login para acessar o sistema.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)

    return decorated_function


# =========================
# Dashboard / Home
# =========================
@app.route("/")
@login_required
def home():
    regiao = request.args.get("regiao", "").strip()

    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    conexao = conectar_oracle()
    total_cabos = 0
    total_contatos = 0
    ranking_cabos = []
    regioes = []

    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return render_template(
            "dashboard.html",
            total_cabos=0,
            total_contatos=0,
            ranking_cabos=[],
            regiao=regiao,
            regioes=[]
        )

    cursor = conexao.cursor()

    try:
        # Regiões disponíveis
        if tipo_acesso == "CABO":
            cursor.execute("""
                SELECT DISTINCT REGIAO_ADMINISTRATIVA
                FROM CABOS_ELEITORAIS
                WHERE ID = :1
                  AND REGIAO_ADMINISTRATIVA IS NOT NULL
                ORDER BY REGIAO_ADMINISTRATIVA
            """, (cabo_sessao,))
        else:
            cursor.execute("""
                SELECT DISTINCT REGIAO_ADMINISTRATIVA
                FROM CABOS_ELEITORAIS
                WHERE REGIAO_ADMINISTRATIVA IS NOT NULL
                ORDER BY REGIAO_ADMINISTRATIVA
            """)
        regioes = [r[0] for r in cursor.fetchall()]

        # Total de lideranças
        if tipo_acesso == "CABO":
            cursor.execute("""
                SELECT COUNT(*)
                FROM CABOS_ELEITORAIS
                WHERE ID = :1
            """, (cabo_sessao,))
            total_cabos = cursor.fetchone()[0]
        else:
            if regiao:
                cursor.execute("""
                    SELECT COUNT(*)
                    FROM CABOS_ELEITORAIS
                    WHERE REGIAO_ADMINISTRATIVA = :1
                """, (regiao,))
            else:
                cursor.execute("SELECT COUNT(*) FROM CABOS_ELEITORAIS")
            total_cabos = cursor.fetchone()[0]

        # Total de apoiadores
        if tipo_acesso == "CABO":
            cursor.execute("""
                SELECT COUNT(*)
                FROM CONTATOS_CAMPANHA
                WHERE CABO_ID = :1
            """, (cabo_sessao,))
            total_contatos = cursor.fetchone()[0]
        else:
            if regiao:
                cursor.execute("""
                    SELECT COUNT(ct.ID)
                    FROM CONTATOS_CAMPANHA ct
                    JOIN CABOS_ELEITORAIS c ON c.ID = ct.CABO_ID
                    WHERE c.REGIAO_ADMINISTRATIVA = :1
                """, (regiao,))
            else:
                cursor.execute("SELECT COUNT(*) FROM CONTATOS_CAMPANHA")
            total_contatos = cursor.fetchone()[0]

        # Ranking / resumo
        sql = """
            SELECT
                c.NOME,
                c.REGIAO_ADMINISTRATIVA,
                COUNT(ct.ID) AS TOTAL_CONTATOS
            FROM CABOS_ELEITORAIS c
            LEFT JOIN CONTATOS_CAMPANHA ct ON ct.CABO_ID = c.ID
            WHERE 1=1
        """
        params = {}

        if regiao:
            sql += " AND c.REGIAO_ADMINISTRATIVA = :regiao"
            params["regiao"] = regiao

        if tipo_acesso == "CABO":
            sql += " AND c.ID = :cabo_sessao"
            params["cabo_sessao"] = cabo_sessao

        sql += """
            GROUP BY c.NOME, c.REGIAO_ADMINISTRATIVA
            ORDER BY TOTAL_CONTATOS DESC, c.NOME
        """

        cursor.execute(sql, params)
        ranking_cabos = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao carregar dashboard: {erro}", "danger")
        print("ERRO DASHBOARD:", erro)
    finally:
        cursor.close()
        conexao.close()

    return render_template(
        "dashboard.html",
        total_cabos=total_cabos,
        total_contatos=total_contatos,
        ranking_cabos=ranking_cabos,
        regiao=regiao,
        regioes=regioes
    )
    
# =========================
# Lideranças / Cabos
# =========================
@app.route("/cadastrar", methods=["GET", "POST"])
@login_required
@perfil_required("ADMIN", "CHEFE_GABINETE", "SECRETARIA")
def cadastrar():
    ...
    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip()
        telefone = request.form.get("telefone", "").strip()
        endereco = request.form.get("endereco", "").strip()
        regiao_administrativa = request.form.get("regiao_administrativa", "").strip()
        data_nascimento = converter_data(request.form.get("data_nascimento", "").strip())
        cep = request.form.get("cep", "").strip()

        if not nome:
            flash("O campo nome é obrigatório.", "warning")
            return render_template("cadastrar.html")

        conexao = conectar_oracle()
        if conexao is None:
            flash("Não foi possível conectar ao banco de dados.", "danger")
            return render_template("cadastrar.html")

        cursor = conexao.cursor()

        try:
            cursor.execute("""
                INSERT INTO CABOS_ELEITORAIS
                (NOME, EMAIL, TELEFONE, ENDERECO, REGIAO_ADMINISTRATIVA, DATA_NASCIMENTO, CEP)
                VALUES (:1, :2, :3, :4, :5, :6, :7)
            """, (nome, email, telefone, endereco, regiao_administrativa, data_nascimento, cep))
            conexao.commit()
            flash("Cabo cadastrado com sucesso.", "success")
            return redirect(url_for("listar"))
        except oracledb.Error as erro:
            print("ERRO AO CADASTRAR CABO:", erro)
            flash(f"Erro ao cadastrar cabo: {erro}", "danger")
        finally:
            cursor.close()
            conexao.close()

    return render_template("cadastrar.html")


@app.route("/listar")
@login_required
def listar():
    termo = request.args.get("busca", "").strip()
    regiao = request.args.get("regiao", "").strip()

    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return render_template("listar.html", registros=[], termo=termo, regiao=regiao, regioes=[])

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT DISTINCT REGIAO_ADMINISTRATIVA
            FROM CABOS_ELEITORAIS
            WHERE REGIAO_ADMINISTRATIVA IS NOT NULL
            ORDER BY REGIAO_ADMINISTRATIVA
        """)
        regioes = [r[0] for r in cursor.fetchall()]

        sql = """
            SELECT
                c.ID,
                c.NOME,
                c.EMAIL,
                c.TELEFONE,
                c.ENDERECO,
                c.REGIAO_ADMINISTRATIVA,
                c.DATA_NASCIMENTO,
                c.CEP,
                COUNT(ct.ID) AS TOTAL_CONTATOS,
                SUM(CASE WHEN ct.CONSENTIU_CONTATO = 'S' THEN 1 ELSE 0 END) AS TOTAL_SIM,
                SUM(CASE WHEN ct.CONSENTIU_CONTATO = 'N' THEN 1 ELSE 0 END) AS TOTAL_NAO
            FROM CABOS_ELEITORAIS c
            LEFT JOIN CONTATOS_CAMPANHA ct ON ct.CABO_ID = c.ID
            WHERE 1=1
        """
        params = {}

        if termo:
            sql += " AND TRIM(UPPER(c.NOME)) LIKE TRIM(UPPER(:busca))"
            params["busca"] = f"%{termo}%"

        if regiao:
            sql += " AND c.REGIAO_ADMINISTRATIVA = :regiao"
            params["regiao"] = regiao

        if tipo_acesso == "CABO":
            sql += " AND c.ID = :cabo_sessao"
            params["cabo_sessao"] = cabo_sessao

        sql += """
            GROUP BY
                c.ID,
                c.NOME,
                c.EMAIL,
                c.TELEFONE,
                c.ENDERECO,
                c.REGIAO_ADMINISTRATIVA,
                c.DATA_NASCIMENTO,
                c.CEP
            ORDER BY c.ID
        """

        cursor.execute(sql, params)
        registros = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao listar dados: {erro}", "danger")
        print("ERRO NA LISTAGEM DE CABOS:", erro)
        registros = []
        regioes = []
    finally:
        cursor.close()
        conexao.close()

    return render_template(
        "listar.html",
        registros=registros,
        termo=termo,
        regiao=regiao,
        regioes=regioes
    )

@app.route("/exportar-excel")
@login_required
def exportar_excel():
    termo = request.args.get("busca", "").strip()
    regiao = request.args.get("regiao", "").strip()

    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar"))

    cursor = conexao.cursor()

    try:
        sql = """
            SELECT
                c.ID,
                c.NOME,
                c.EMAIL,
                c.TELEFONE,
                c.ENDERECO,
                c.REGIAO_ADMINISTRATIVA,
                c.DATA_NASCIMENTO,
                c.CEP,
                COUNT(ct.ID) AS TOTAL_CONTATOS
            FROM CABOS_ELEITORAIS c
            LEFT JOIN CONTATOS_CAMPANHA ct ON ct.CABO_ID = c.ID
            WHERE 1=1
        """
        params = {}

        if termo:
            sql += " AND TRIM(UPPER(c.NOME)) LIKE TRIM(UPPER(:busca))"
            params["busca"] = f"%{termo}%"

        if regiao:
            sql += " AND c.REGIAO_ADMINISTRATIVA = :regiao"
            params["regiao"] = regiao

        if tipo_acesso == "CABO":
            sql += " AND c.ID = :cabo_sessao"
            params["cabo_sessao"] = cabo_sessao

        sql += """
            GROUP BY
                c.ID,
                c.NOME,
                c.EMAIL,
                c.TELEFONE,
                c.ENDERECO,
                c.REGIAO_ADMINISTRATIVA,
                c.DATA_NASCIMENTO,
                c.CEP
            ORDER BY c.ID
        """

        cursor.execute(sql, params)
        registros = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao exportar dados: {erro}", "danger")
        print("ERRO AO EXPORTAR CABOS:", erro)
        return redirect(url_for("listar"))
    finally:
        cursor.close()
        conexao.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Lideranças"

    sheet.append([
        "ID",
        "Nome",
        "Email",
        "Telefone",
        "Endereço",
        "Região Administrativa",
        "Data de Nascimento",
        "CEP",
        "Total de Contatos"
    ])

    for registro in registros:
        linha = list(registro)
        if linha[6]:
            linha[6] = linha[6].strftime("%d/%m/%Y")
        sheet.append(linha)

    arquivo_excel = BytesIO()
    workbook.save(arquivo_excel)
    arquivo_excel.seek(0)

    return send_file(
        arquivo_excel,
        as_attachment=True,
        download_name="liderancas.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
@app.route("/editar/<int:id>", methods=["GET", "POST"])
@login_required
def editar(id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar"))

    cursor = conexao.cursor()

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip()
        telefone = request.form.get("telefone", "").strip()
        endereco = request.form.get("endereco", "").strip()
        regiao_administrativa = request.form.get("regiao_administrativa", "").strip()
        data_nascimento = converter_data(request.form.get("data_nascimento", "").strip())
        cep = request.form.get("cep", "").strip()

        try:
            cursor.execute("""
                UPDATE CABOS_ELEITORAIS
                SET NOME = :1,
                    EMAIL = :2,
                    TELEFONE = :3,
                    ENDERECO = :4,
                    REGIAO_ADMINISTRATIVA = :5,
                    DATA_NASCIMENTO = :6,
                    CEP = :7
                WHERE ID = :8
            """, (nome, email, telefone, endereco, regiao_administrativa, data_nascimento, cep, id))
            conexao.commit()
            flash("Registro atualizado com sucesso.", "success")
            return redirect(url_for("listar"))
        except oracledb.Error as erro:
            print("ERRO AO EDITAR CABO:", erro)
            flash(f"Erro ao atualizar cabo: {erro}", "danger")
        finally:
            cursor.close()
            conexao.close()

    try:
        cursor.execute("""
            SELECT ID, NOME, EMAIL, TELEFONE, ENDERECO, REGIAO_ADMINISTRATIVA, DATA_NASCIMENTO, CEP
            FROM CABOS_ELEITORAIS
            WHERE ID = :1
        """, (id,))
        registro = cursor.fetchone()
    except oracledb.Error as erro:
        print("ERRO AO BUSCAR CABO:", erro)
        flash(f"Erro ao buscar registro: {erro}", "danger")
        registro = None
    finally:
        cursor.close()
        conexao.close()

    if not registro:
        flash("Registro não encontrado.", "warning")
        return redirect(url_for("listar"))

    return render_template("editar.html", registro=registro)


@app.route("/excluir/<int:id>")
@login_required
@perfil_required("ADMIN", "CHEFE_GABINETE")
def excluir(id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT COUNT(*)
            FROM CONTATOS_CAMPANHA
            WHERE CABO_ID = :1
        """, (id,))
        total_contatos = cursor.fetchone()[0]

        cursor.execute("""
            DELETE FROM CONVITES_CONTATO
            WHERE CABO_ID = :1
              AND STATUS = 'CANCELADO'
        """, (id,))

        cursor.execute("""
            SELECT COUNT(*)
            FROM CONVITES_CONTATO
            WHERE CABO_ID = :1
        """, (id,))
        total_convites = cursor.fetchone()[0]

        if total_contatos > 0 or total_convites > 0:
            conexao.rollback()
            flash(
                f"Não é possível excluir esta liderança. "
                f"Existem {total_contatos} apoiador(es) e "
                f"{total_convites} convite(s) ativo(s)/histórico(s) vinculados a ela.",
                "warning"
            )
            return redirect(url_for("listar"))

        cursor.execute("""
            DELETE FROM CABOS_ELEITORAIS
            WHERE ID = :1
        """, (id,))

        conexao.commit()
        flash("Liderança excluída com sucesso.", "success")

    except oracledb.Error as erro:
        conexao.rollback()
        flash(f"Erro ao excluir liderança: {erro}", "danger")
        print("ERRO AO EXCLUIR LIDERANÇA:", erro)

    finally:
        cursor.close()
        conexao.close()

    return redirect(url_for("listar"))

# =========================
# Contatos / Apoiadores
# =========================
@app.route("/cadastrar-contato", methods=["GET", "POST"])
@login_required
@perfil_required("ADMIN", "CHEFE_GABINETE", "SECRETARIA")
def cadastrar_contato():
    ...
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("home"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT ID, NOME
            FROM CABOS_ELEITORAIS
            ORDER BY NOME
        """)
        cabos = cursor.fetchall()
    except oracledb.Error as erro:
        cursor.close()
        conexao.close()
        flash(f"Erro ao carregar cabos: {erro}", "danger")
        return redirect(url_for("home"))

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip().lower()
        telefone = request.form.get("telefone", "").strip()
        endereco = request.form.get("endereco", "").strip()
        regiao = request.form.get("regiao_administrativa", "").strip()
        cabo_id = request.form.get("cabo_id", "").strip()
        consentiu = request.form.get("consentiu_contato", "N")
        observacao = request.form.get("observacao", "").strip()
        data_nascimento = converter_data(request.form.get("data_nascimento", "").strip())
        cep = request.form.get("cep", "").strip()

        if not nome or not cabo_id:
            flash("Nome e cabo responsável são obrigatórios.", "warning")
            cursor.close()
            conexao.close()
            return render_template("cadastrar_contato.html", cabos=cabos)

        try:
            cursor.execute("""
                INSERT INTO CONTATOS_CAMPANHA
                (NOME, EMAIL, TELEFONE, ENDERECO, REGIAO_ADMINISTRATIVA, CABO_ID,
                CONSENTIU_CONTATO, OBSERVACAO, DATA_NASCIMENTO, CEP)
                VALUES (:1, :2, :3, :4, :5, :6, :7, :8, :9, :10)
            """, (
                nome,
                email,
                telefone,
                endereco,
                regiao,
                int(cabo_id),
                consentiu,
                observacao,
                data_nascimento,
                cep
            ))
            conexao.commit()
            flash("Contato cadastrado com sucesso.", "success")
            return redirect(url_for("listar_contatos"))
        except oracledb.Error as erro:
            flash(f"Erro ao cadastrar contato: {erro}", "danger")
        finally:
            cursor.close()
            conexao.close()

    cursor.close()
    conexao.close()
    return render_template("cadastrar_contato.html", cabos=cabos)


@app.route("/listar-contatos")
@login_required
def listar_contatos():
    termo = request.args.get("busca", "").strip()
    cabo_id = request.args.get("cabo_id", "").strip()
    regiao = request.args.get("regiao", "").strip()

    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return render_template(
            "listar_contatos.html",
            registros=[],
            termo=termo,
            cabo_id=cabo_id,
            regiao=regiao,
            cabos=[],
            regioes=[]
        )

    cursor = conexao.cursor()

    try:
        if tipo_acesso == "CABO":
            cursor.execute("""
                SELECT ID, NOME
                FROM CABOS_ELEITORAIS
                WHERE ID = :1
                ORDER BY NOME
            """, (cabo_sessao,))
        else:
            cursor.execute("""
                SELECT ID, NOME
                FROM CABOS_ELEITORAIS
                ORDER BY NOME
            """)
        cabos = cursor.fetchall()

        cursor.execute("""
            SELECT DISTINCT REGIAO_ADMINISTRATIVA
            FROM CONTATOS_CAMPANHA
            WHERE REGIAO_ADMINISTRATIVA IS NOT NULL
            ORDER BY REGIAO_ADMINISTRATIVA
        """)
        regioes = [r[0] for r in cursor.fetchall()]

        sql = """
            SELECT
                ct.ID,
                ct.NOME,
                ct.EMAIL,
                ct.TELEFONE,
                ct.REGIAO_ADMINISTRATIVA,
                c.NOME AS CABO_NOME,
                ct.DATA_NASCIMENTO,
                ct.CEP,
                ct.CONSENTIU_CONTATO
            FROM CONTATOS_CAMPANHA ct
            JOIN CABOS_ELEITORAIS c ON c.ID = ct.CABO_ID
            WHERE 1=1
        """
        params = {}

        if termo:
            sql += " AND UPPER(ct.NOME) LIKE :busca"
            params["busca"] = f"%{termo.upper()}%"

        if cabo_id:
            sql += " AND ct.CABO_ID = :cabo_id"
            params["cabo_id"] = int(cabo_id)

        if regiao:
            sql += " AND ct.REGIAO_ADMINISTRATIVA = :regiao"
            params["regiao"] = regiao

        if tipo_acesso == "CABO":
            sql += " AND ct.CABO_ID = :cabo_sessao"
            params["cabo_sessao"] = cabo_sessao

        sql += " ORDER BY ct.ID"

        cursor.execute(sql, params)
        registros = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao listar contatos: {erro}", "danger")
        print("ERRO LISTAR CONTATOS:", erro)
        registros = []
        cabos = []
        regioes = []
    finally:
        cursor.close()
        conexao.close()

    return render_template(
        "listar_contatos.html",
        registros=registros,
        termo=termo,
        cabo_id=cabo_id,
        regiao=regiao,
        cabos=cabos,
        regioes=regioes
    )
    
@app.route("/exportar-contatos-excel")
@login_required
def exportar_contatos_excel():
    termo = request.args.get("busca", "").strip()
    cabo_id = request.args.get("cabo_id", "").strip()
    regiao = request.args.get("regiao", "").strip()

    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_contatos"))

    cursor = conexao.cursor()

    try:
        sql = """
            SELECT
                ct.ID,
                ct.NOME,
                ct.EMAIL,
                ct.TELEFONE,
                ct.ENDERECO,
                ct.REGIAO_ADMINISTRATIVA,
                c.NOME AS CABO_NOME,
                ct.CONSENTIU_CONTATO,
                ct.OBSERVACAO,
                ct.DATA_CADASTRO,
                ct.DATA_NASCIMENTO,
                ct.CEP
            FROM CONTATOS_CAMPANHA ct
            JOIN CABOS_ELEITORAIS c ON c.ID = ct.CABO_ID
            WHERE 1=1
        """
        params = {}

        if termo:
            sql += " AND UPPER(ct.NOME) LIKE :busca"
            params["busca"] = f"%{termo.upper()}%"

        if cabo_id:
            sql += " AND ct.CABO_ID = :cabo_id"
            params["cabo_id"] = int(cabo_id)

        if regiao:
            sql += " AND ct.REGIAO_ADMINISTRATIVA = :regiao"
            params["regiao"] = regiao

        if tipo_acesso == "CABO":
            sql += " AND ct.CABO_ID = :cabo_sessao"
            params["cabo_sessao"] = cabo_sessao

        sql += " ORDER BY ct.ID"

        cursor.execute(sql, params)
        registros = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao exportar contatos: {erro}", "danger")
        print("ERRO AO EXPORTAR CONTATOS:", erro)
        return redirect(url_for("listar_contatos"))
    finally:
        cursor.close()
        conexao.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Apoiadores"

    sheet.append([
        "ID",
        "Nome",
        "Email",
        "Telefone",
        "Endereço",
        "Região Administrativa",
        "Liderança",
        "Consentiu Contato",
        "Observação",
        "Data Cadastro",
        "Data Nascimento",
        "CEP"
    ])

    for registro in registros:
        linha = list(registro)

        if linha[9]:
            linha[9] = linha[9].strftime("%d/%m/%Y")

        if linha[10]:
            linha[10] = linha[10].strftime("%d/%m/%Y")

        sheet.append(linha)

    arquivo_excel = BytesIO()
    workbook.save(arquivo_excel)
    arquivo_excel.seek(0)

    return send_file(
        arquivo_excel,
        as_attachment=True,
        download_name="apoiadores.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/editar-contato/<int:id>", methods=["GET", "POST"])
@login_required
def editar_contato(id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_contatos"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT ID, NOME
            FROM CABOS_ELEITORAIS
            ORDER BY NOME
        """)
        cabos = cursor.fetchall()

        if request.method == "POST":
            nome = request.form.get("nome", "").strip()
            email = request.form.get("email", "").strip().lower()
            telefone = request.form.get("telefone", "").strip()
            endereco = request.form.get("endereco", "").strip()
            regiao = request.form.get("regiao_administrativa", "").strip()
            cabo_id = request.form.get("cabo_id", "").strip()
            consentiu = request.form.get("consentiu_contato", "N")
            observacao = request.form.get("observacao", "").strip()
            data_nascimento = converter_data(request.form.get("data_nascimento", "").strip())
            cep = request.form.get("cep", "").strip()

            if not nome or not cabo_id:
                flash("Nome e liderança responsável são obrigatórios.", "warning")
            else:
                # validação opcional para não duplicar e-mail em outro contato
                if email:
                    cursor.execute("""
                        SELECT ID
                        FROM CONTATOS_CAMPANHA
                        WHERE LOWER(EMAIL) = LOWER(:1)
                          AND ID <> :2
                    """, (email, id))
                    contato_email_existente = cursor.fetchone()

                    if contato_email_existente:
                        flash("Já existe outro contato cadastrado com esse e-mail.", "warning")
                        cursor.execute("""
                            SELECT
                                ID,
                                NOME,
                                EMAIL,
                                TELEFONE,
                                ENDERECO,
                                REGIAO_ADMINISTRATIVA,
                                CABO_ID,
                                CONSENTIU_CONTATO,
                                OBSERVACAO,
                                DATA_NASCIMENTO,
                                CEP
                            FROM CONTATOS_CAMPANHA
                            WHERE ID = :1
                        """, (id,))
                        contato = cursor.fetchone()
                        return render_template("editar_contato.html", contato=contato, cabos=cabos)

                cursor.execute("""
                    UPDATE CONTATOS_CAMPANHA
                    SET NOME = :1,
                        EMAIL = :2,
                        TELEFONE = :3,
                        ENDERECO = :4,
                        REGIAO_ADMINISTRATIVA = :5,
                        CABO_ID = :6,
                        CONSENTIU_CONTATO = :7,
                        OBSERVACAO = :8,
                        DATA_NASCIMENTO = :9,
                        CEP = :10
                    WHERE ID = :11
                """, (
                    nome,
                    email,
                    telefone,
                    endereco,
                    regiao,
                    int(cabo_id),
                    consentiu,
                    observacao,
                    data_nascimento,
                    cep,
                    id
                ))
                conexao.commit()
                flash("Contato atualizado com sucesso.", "success")
                return redirect(url_for("listar_contatos"))

        cursor.execute("""
            SELECT
                ID,
                NOME,
                EMAIL,
                TELEFONE,
                ENDERECO,
                REGIAO_ADMINISTRATIVA,
                CABO_ID,
                CONSENTIU_CONTATO,
                OBSERVACAO,
                DATA_NASCIMENTO,
                CEP
            FROM CONTATOS_CAMPANHA
            WHERE ID = :1
        """, (id,))
        contato = cursor.fetchone()

        if not contato:
            flash("Contato não encontrado.", "warning")
            return redirect(url_for("listar_contatos"))

    except oracledb.Error as erro:
        print("ERRO AO EDITAR CONTATO:", erro)
        flash(f"Erro ao carregar/atualizar contato: {erro}", "danger")
        return redirect(url_for("listar_contatos"))
    finally:
        cursor.close()
        conexao.close()

    return render_template("editar_contato.html", contato=contato, cabos=cabos)


@app.route("/excluir-contato/<int:id>")
@login_required
def excluir_contato(id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_contatos"))

    cursor = conexao.cursor()

    try:
        cursor.execute("DELETE FROM CONTATOS_CAMPANHA WHERE ID = :1", (id,))
        conexao.commit()
        flash("Contato excluído com sucesso.", "success")
    except oracledb.Error as erro:
        flash(f"Erro ao excluir contato: {erro}", "danger")
    finally:
        cursor.close()
        conexao.close()

    return redirect(url_for("listar_contatos"))


@app.route("/cabo/<int:cabo_id>/contatos")
@login_required
def contatos_do_cabo(cabo_id):
    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    if tipo_acesso == "CABO" and cabo_sessao != cabo_id:
        flash("Você só pode acessar os seus próprios contatos.", "danger")
        return redirect(url_for("home"))

   
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT ID, NOME, REGIAO_ADMINISTRATIVA
            FROM CABOS_ELEITORAIS
            WHERE ID = :1
        """, (cabo_id,))
        cabo = cursor.fetchone()

        if not cabo:
            flash("Liderança não encontrada.", "warning")
            return redirect(url_for("listar"))

        cursor.execute("""
            SELECT
                ID,                    -- item[0]
                NOME,                  -- item[1]
                TELEFONE,              -- item[2]
                ENDERECO,              -- item[3]
                DATA_NASCIMENTO,       -- item[4]
                CEP,                   -- item[5]
                REGIAO_ADMINISTRATIVA, -- item[6]
                CONSENTIU_CONTATO,     -- item[7]
                OBSERVACAO,            -- item[8]
                DATA_CADASTRO          -- item[9]
            FROM CONTATOS_CAMPANHA
            WHERE CABO_ID = :1
            ORDER BY NOME
        """, (cabo_id,))
        contatos = cursor.fetchall()

        total_contatos = len(contatos)

    except oracledb.Error as erro:
        flash(f"Erro ao carregar contatos da liderança: {erro}", "danger")
        print("ERRO CONTATOS DO LÍDER:", erro)
        return redirect(url_for("listar"))
    finally:
        cursor.close()
        conexao.close()

    return render_template(
        "contatos_do_cabo.html",
        cabo=cabo,
        contatos=contatos,
        total_contatos=total_contatos
    )

@app.route("/cabo/<int:cabo_id>/exportar-excel")
@login_required
def exportar_contatos_do_cabo(cabo_id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT NOME
            FROM CABOS_ELEITORAIS
            WHERE ID = :1
        """, (cabo_id,))
        cabo = cursor.fetchone()

        if not cabo:
            flash("Cabo não encontrado.", "warning")
            return redirect(url_for("listar"))

        nome_cabo = cabo[0]

        cursor.execute("""
            SELECT
                ID,
                NOME,
                TELEFONE,
                ENDERECO,
                REGIAO_ADMINISTRATIVA,
                CONSENTIU_CONTATO,
                OBSERVACAO,
                DATA_CADASTRO,
                DATA_NASCIMENTO,
                CEP
            FROM CONTATOS_CAMPANHA
            WHERE CABO_ID = :1
            ORDER BY ID
        """, (cabo_id,))
        contatos = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao exportar contatos: {erro}", "danger")
        print("ERRO AO EXPORTAR CONTATOS DO CABO:", erro)
        return redirect(url_for("listar"))
    finally:
        cursor.close()
        conexao.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Contatos do Cabo"

    sheet.append([
        "ID",
        "Nome",
        "Telefone",
        "Endereço",
        "Região Administrativa",
        "Consentiu Contato",
        "Observação",
        "Data Cadastro",
        "Data Nascimento",
        "CEP"
    ])

    for contato in contatos:
        linha = list(contato)
        if linha[7]:
            linha[7] = linha[7].strftime("%d/%m/%Y")
        if linha[8]:
            linha[8] = linha[8].strftime("%d/%m/%Y")
        sheet.append(linha)

    arquivo_excel = BytesIO()
    workbook.save(arquivo_excel)
    arquivo_excel.seek(0)

    nome_arquivo = f"contatos_{nome_cabo.replace(' ', '_').lower()}.xlsx"

    return send_file(
        arquivo_excel,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# Relatórios
# =========================
@app.route("/relatorio-geral")
@login_required
@perfil_required("ADMIN", "DEPUTADO", "CHEFE_GABINETE")
def relatorio_geral():
    ...
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("home"))

    cursor = conexao.cursor()
    total_cabos = 0
    total_contatos = 0
    contatos_por_regiao = []
    ranking_cabos = []

    try:
        cursor.execute("SELECT COUNT(*) FROM CABOS_ELEITORAIS")
        total_cabos = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM CONTATOS_CAMPANHA")
        total_contatos = cursor.fetchone()[0]

        cursor.execute("""
            SELECT
                NVL(REGIAO_ADMINISTRATIVA, 'Não informada') AS REGIAO,
                COUNT(*) AS TOTAL
            FROM CONTATOS_CAMPANHA
            GROUP BY NVL(REGIAO_ADMINISTRATIVA, 'Não informada')
            ORDER BY TOTAL DESC, REGIAO
        """)
        contatos_por_regiao = cursor.fetchall()

        cursor.execute("""
            SELECT
                c.NOME,
                NVL(c.REGIAO_ADMINISTRATIVA, 'Não informada') AS REGIAO,
                COUNT(ct.ID) AS TOTAL_CONTATOS
            FROM CABOS_ELEITORAIS c
            LEFT JOIN CONTATOS_CAMPANHA ct ON ct.CABO_ID = c.ID
            GROUP BY c.NOME, NVL(c.REGIAO_ADMINISTRATIVA, 'Não informada')
            ORDER BY TOTAL_CONTATOS DESC, c.NOME
        """)
        ranking_cabos = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao carregar relatório: {erro}", "danger")
        return redirect(url_for("home"))
    finally:
        cursor.close()
        conexao.close()

    regioes_labels = [item[0] for item in contatos_por_regiao]
    regioes_valores = [item[1] for item in contatos_por_regiao]
    cabos_labels = [item[0] for item in ranking_cabos]
    cabos_valores = [item[2] for item in ranking_cabos]

    return render_template(
        "relatorio_geral.html",
        total_cabos=total_cabos,
        total_contatos=total_contatos,
        contatos_por_regiao=contatos_por_regiao,
        ranking_cabos=ranking_cabos,
        regioes_labels=regioes_labels,
        regioes_valores=regioes_valores,
        cabos_labels=cabos_labels,
        cabos_valores=cabos_valores
    )


@app.route("/exportar-relatorio-geral-excel")
@login_required
def exportar_relatorio_geral_excel():
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("relatorio_geral"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT
                c.NOME,
                NVL(c.REGIAO_ADMINISTRATIVA, 'Não informada') AS REGIAO,
                COUNT(ct.ID) AS TOTAL_CONTATOS
            FROM CABOS_ELEITORAIS c
            LEFT JOIN CONTATOS_CAMPANHA ct ON ct.CABO_ID = c.ID
            GROUP BY c.NOME, NVL(c.REGIAO_ADMINISTRATIVA, 'Não informada')
            ORDER BY TOTAL_CONTATOS DESC, c.NOME
        """)
        ranking_cabos = cursor.fetchall()

        cursor.execute("""
            SELECT
                NVL(REGIAO_ADMINISTRATIVA, 'Não informada') AS REGIAO,
                COUNT(*) AS TOTAL
            FROM CONTATOS_CAMPANHA
            GROUP BY NVL(REGIAO_ADMINISTRATIVA, 'Não informada')
            ORDER BY TOTAL DESC, REGIAO
        """)
        contatos_por_regiao = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao exportar relatório: {erro}", "danger")
        return redirect(url_for("relatorio_geral"))
    finally:
        cursor.close()
        conexao.close()

    workbook = Workbook()

    aba1 = workbook.active
    aba1.title = "Ranking Cabos"
    aba1.append(["Cabo", "Região", "Total de Contatos"])
    for linha in ranking_cabos:
        aba1.append(list(linha))

    aba2 = workbook.create_sheet(title="Contatos por Região")
    aba2.append(["Região", "Total de Contatos"])
    for linha in contatos_por_regiao:
        aba2.append(list(linha))

    arquivo_excel = BytesIO()
    workbook.save(arquivo_excel)
    arquivo_excel.seek(0)

    return send_file(
        arquivo_excel,
        as_attachment=True,
        download_name="relatorio_geral.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# Login
# =========================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("usuario", "").strip().lower()
        senha = request.form.get("senha", "").strip()

        conexao = conectar_oracle()
        if conexao is None:
            flash("Não foi possível conectar ao banco de dados.", "danger")
            return render_template("login.html")

        cursor = conexao.cursor()

        try:
            cursor.execute("""
                SELECT
                    ID,
                    NOME,
                    EMAIL,
                    SENHA_HASH,
                    TIPO_ACESSO,
                    CABO_ID,
                    ATIVO
                FROM USUARIOS_SISTEMA
                WHERE LOWER(EMAIL) = LOWER(:1)
            """, (email,))
            usuario = cursor.fetchone()

            if not usuario:
                flash("Usuário ou senha inválidos.", "danger")
                return render_template("login.html")

            usuario_id = usuario[0]
            nome = usuario[1]
            email_db = usuario[2]
            senha_hash = usuario[3]
            tipo_acesso = usuario[4]
            cabo_id = usuario[5]
            ativo = usuario[6]

            if ativo != "S":
                flash("Usuário inativo. Procure o administrador.", "warning")
                return render_template("login.html")

            if not check_password_hash(senha_hash, senha):
                flash("Usuário ou senha inválidos.", "danger")
                return render_template("login.html")

            session["usuario_logado"] = True
            session["usuario_id"] = usuario_id
            session["usuario_nome"] = nome
            session["usuario_email"] = email_db
            session["tipo_acesso"] = tipo_acesso
            session["cabo_id"] = cabo_id

            flash("Login realizado com sucesso.", "success")
            return redirect(url_for("home"))

        except oracledb.Error as erro:
            flash(f"Erro ao realizar login: {erro}", "danger")
            print("ERRO LOGIN:", erro)
        finally:
            cursor.close()
            conexao.close()

    return render_template("login.html")

def perfil_required(*tipos_permitidos):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not session.get("usuario_logado"):
                flash("Faça login para acessar o sistema.", "warning")
                return redirect(url_for("login"))

            tipo_usuario = session.get("tipo_acesso")

            if tipo_usuario not in tipos_permitidos:
                flash("Você não tem permissão para acessar esta área.", "danger")
                return redirect(url_for("home"))

            return f(*args, **kwargs)
        return decorated_function
    return decorator

@app.route("/logout")
def logout():
    session.clear()
    flash("Logout realizado com sucesso.", "success")
    return redirect(url_for("login"))

#Rotas de gerenciamento de usuários (apenas para ADMIN)

@app.route("/cadastrar-usuario", methods=["GET", "POST"])
@login_required
@perfil_required("ADMIN")
def cadastrar_usuario():
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("home"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT ID, NOME
            FROM CABOS_ELEITORAIS
            ORDER BY NOME
        """)
        cabos = cursor.fetchall()

        if request.method == "POST":
            nome = request.form.get("nome", "").strip()
            email = request.form.get("email", "").strip().lower()
            senha = request.form.get("senha", "").strip()
            tipo_acesso = request.form.get("tipo_acesso", "").strip()
            cabo_id = request.form.get("cabo_id", "").strip()
            ativo = request.form.get("ativo", "S").strip()

            if not nome or not email or not senha or not tipo_acesso:
                flash("Preencha os campos obrigatórios.", "warning")
                return render_template("cadastrar_usuario.html", cabos=cabos)

            cursor.execute("""
                SELECT ID
                FROM USUARIOS_SISTEMA
                WHERE LOWER(EMAIL) = LOWER(:1)
            """, (email,))
            usuario_existente = cursor.fetchone()

            if usuario_existente:
                flash("Já existe um usuário cadastrado com esse e-mail.", "warning")
                return render_template("cadastrar_usuario.html", cabos=cabos)

            if tipo_acesso == "CABO" and not cabo_id:
                flash("Para usuário do tipo CABO, selecione a liderança.", "warning")
                return render_template("cadastrar_usuario.html", cabos=cabos)

            senha_hash = generate_password_hash(senha)

            cursor.execute("""
                INSERT INTO USUARIOS_SISTEMA
                    (NOME, EMAIL, SENHA_HASH, TIPO_ACESSO, CABO_ID, ATIVO)
                VALUES
                    (:1, :2, :3, :4, :5, :6)
            """, (
                nome,
                email,
                senha_hash,
                tipo_acesso,
                int(cabo_id) if cabo_id else None,
                ativo
            ))
            conexao.commit()

            flash("Usuário cadastrado com sucesso.", "success")
            return redirect(url_for("listar_usuarios"))

    except oracledb.Error as erro:
        flash(f"Erro ao cadastrar usuário: {erro}", "danger")
        print("ERRO CADASTRAR USUÁRIO:", erro)
        return redirect(url_for("home"))
    finally:
        cursor.close()
        conexao.close()

    return render_template("cadastrar_usuario.html", cabos=cabos)

@app.route("/listar-usuarios")
@login_required
@perfil_required("ADMIN")
def listar_usuarios():
    ...
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("home"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT
                u.ID,
                u.NOME,
                u.EMAIL,
                u.TIPO_ACESSO,
                u.CABO_ID,
                c.NOME AS NOME_CABO,
                u.ATIVO,
                u.DATA_CADASTRO
            FROM USUARIOS_SISTEMA u
            LEFT JOIN CABOS_ELEITORAIS c ON c.ID = u.CABO_ID
            ORDER BY u.ID
        """)
        usuarios = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao listar usuários: {erro}", "danger")
        print("ERRO LISTAR USUÁRIOS:", erro)
        return redirect(url_for("home"))
    finally:
        cursor.close()
        conexao.close()

    return render_template("listar_usuarios.html", usuarios=usuarios)

@app.route("/editar-usuario/<int:id>", methods=["GET", "POST"])
@login_required
@perfil_required("ADMIN")
def editar_usuario(id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_usuarios"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT ID, NOME
            FROM CABOS_ELEITORAIS
            ORDER BY NOME
        """)
        cabos = cursor.fetchall()

        if request.method == "POST":
            nome = request.form.get("nome", "").strip()
            email = request.form.get("email", "").strip().lower()
            tipo_acesso = request.form.get("tipo_acesso", "").strip()
            cabo_id = request.form.get("cabo_id", "").strip()
            ativo = request.form.get("ativo", "S").strip()
            nova_senha = request.form.get("senha", "").strip()

            if not nome or not email or not tipo_acesso:
                flash("Preencha os campos obrigatórios.", "warning")
            else:
                cursor.execute("""
                    SELECT ID
                    FROM USUARIOS_SISTEMA
                    WHERE LOWER(EMAIL) = LOWER(:1)
                      AND ID <> :2
                """, (email, id))
                email_existente = cursor.fetchone()

                if email_existente:
                    flash("Já existe outro usuário com esse e-mail.", "warning")
                    return render_template("editar_usuario.html", usuario=usuario, cabos=cabos)
                else:
                    if tipo_acesso == "CABO" and not cabo_id:
                        flash("Para usuário do tipo CABO, selecione a liderança.", "warning")
                    else:
                        if nova_senha:
                            senha_hash = generate_password_hash(nova_senha)
                            cursor.execute("""
                                UPDATE USUARIOS_SISTEMA
                                SET NOME = :1,
                                    EMAIL = :2,
                                    SENHA_HASH = :3,
                                    TIPO_ACESSO = :4,
                                    CABO_ID = :5,
                                    ATIVO = :6
                                WHERE ID = :7
                            """, (
                                nome,
                                email,
                                senha_hash,
                                tipo_acesso,
                                int(cabo_id) if cabo_id else None,
                                ativo,
                                id
                            ))
                        else:
                            cursor.execute("""
                                UPDATE USUARIOS_SISTEMA
                                SET NOME = :1,
                                    EMAIL = :2,
                                    TIPO_ACESSO = :3,
                                    CABO_ID = :4,
                                    ATIVO = :5
                                WHERE ID = :6
                            """, (
                                nome,
                                email,
                                tipo_acesso,
                                int(cabo_id) if cabo_id else None,
                                ativo,
                                id
                            ))

                        conexao.commit()
                        flash("Usuário atualizado com sucesso.", "success")
                        return redirect(url_for("listar_usuarios"))

        cursor.execute("""
            SELECT
                ID,
                NOME,
                EMAIL,
                TIPO_ACESSO,
                CABO_ID,
                ATIVO
            FROM USUARIOS_SISTEMA
            WHERE ID = :1
        """, (id,))
        usuario = cursor.fetchone()

        if not usuario:
            flash("Usuário não encontrado.", "warning")
            return redirect(url_for("listar_usuarios"))

    except oracledb.Error as erro:
        flash(f"Erro ao editar usuário: {erro}", "danger")
        print("ERRO EDITAR USUÁRIO:", erro)
        return redirect(url_for("listar_usuarios"))
    finally:
        cursor.close()
        conexao.close()

    return render_template("editar_usuario.html", usuario=usuario, cabos=cabos)

@app.route("/inativar-usuario/<int:id>")
@login_required
@perfil_required("ADMIN")
def inativar_usuario(id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_usuarios"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            UPDATE USUARIOS_SISTEMA
            SET ATIVO = 'N'
            WHERE ID = :1
        """, (id,))
        conexao.commit()
        flash("Usuário inativado com sucesso.", "success")
    except oracledb.Error as erro:
        flash(f"Erro ao inativar usuário: {erro}", "danger")
        print("ERRO INATIVAR USUÁRIO:", erro)
    finally:
        cursor.close()
        conexao.close()

    return redirect(url_for("listar_usuarios"))



# =========================
# Convites
# =========================
@app.route("/enviar-convite/<int:cabo_id>", methods=["GET", "POST"])
@login_required
def enviar_convite(cabo_id):
    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    if tipo_acesso == "CABO" and cabo_sessao != cabo_id:
        flash("Você só pode enviar convites da sua própria liderança.", "danger")
        return redirect(url_for("home"))

    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar"))

    cursor = conexao.cursor()
    token = None

    try:
        cursor.execute("""
            SELECT ID, NOME, EMAIL
            FROM CABOS_ELEITORAIS
            WHERE ID = :1
        """, (cabo_id,))
        cabo = cursor.fetchone()

        if not cabo:
            flash("Liderança não encontrada.", "warning")
            return redirect(url_for("listar"))

        if request.method == "POST":
            email = request.form.get("email", "").strip().lower()

            if not email:
                flash("Informe o e-mail do contato.", "warning")
                return render_template("enviar_convite.html", cabo=cabo)

            # 1) Verifica se já existe cadastro com esse e-mail no sistema
            cursor.execute("""
                SELECT COUNT(*)
                FROM USER_TAB_COLUMNS
                WHERE TABLE_NAME = 'CONTATOS_CAMPANHA'
                  AND COLUMN_NAME = 'EMAIL'
            """)
            tem_coluna_email = cursor.fetchone()[0] > 0

            if tem_coluna_email:
                cursor.execute("""
                    SELECT ID
                    FROM CONTATOS_CAMPANHA
                    WHERE LOWER(EMAIL) = LOWER(:1)
                """, (email,))
                contato_existente = cursor.fetchone()

                if contato_existente:
                    flash("Já existe um cadastro no sistema com esse e-mail.", "warning")
                    return render_template("enviar_convite.html", cabo=cabo)

            # 2) Verifica se já existe convite pendente para esse e-mail em qualquer liderança
            cursor.execute("""
                SELECT cc.ID, c.NOME
                FROM CONVITES_CONTATO cc
                JOIN CABOS_ELEITORAIS c ON c.ID = cc.CABO_ID
                WHERE LOWER(cc.EMAIL) = LOWER(:1)
                  AND cc.STATUS = 'PENDENTE'
            """, (email,))
            convite_existente = cursor.fetchone()

            if convite_existente:
                flash("Já existe um convite pendente para este e-mail no sistema.", "warning")
                return render_template("enviar_convite.html", cabo=cabo)

            token = gerar_token_convite()

            cursor.execute("""
                INSERT INTO CONVITES_CONTATO
                    (EMAIL, CABO_ID, TOKEN, STATUS, DATA_EXPIRACAO, STATUS_ENVIO, ERRO_ENVIO)
                VALUES
                    (:1, :2, :3, 'PENDENTE', SYSDATE + 7, 'PENDENTE', NULL)
            """, (email, cabo_id, token))
            conexao.commit()

            link_convite = f"{BASE_URL}/convite/{token}"

            recusados = enviar_email_convite(email, cabo[1], link_convite)

            if recusados:
                cursor.execute("""
                    UPDATE CONVITES_CONTATO
                    SET STATUS_ENVIO = 'RECUSADO',
                        ERRO_ENVIO = :1
                    WHERE TOKEN = :2
                """, (str(recusados)[:500], token))
                conexao.commit()
                flash("O servidor recusou o destinatário.", "warning")
            else:
                cursor.execute("""
                    UPDATE CONVITES_CONTATO
                    SET STATUS_ENVIO = 'ENVIADO',
                        ERRO_ENVIO = NULL
                    WHERE TOKEN = :1
                """, (token,))
                conexao.commit()
                flash("Convite enviado com sucesso.", "success")

            return redirect(url_for("listar"))

    except oracledb.Error as erro:
        if token:
            try:
                cursor.execute("""
                    UPDATE CONVITES_CONTATO
                    SET STATUS_ENVIO = 'ERRO_BANCO',
                        ERRO_ENVIO = :1
                    WHERE TOKEN = :2
                """, (str(erro)[:500], token))
                conexao.commit()
            except Exception:
                pass

        flash(f"Erro ao processar convite: {erro}", "danger")
        print("ERRO AO ENVIAR CONVITE:", erro)
        return redirect(url_for("listar"))

    except Exception as erro:
        if token:
            try:
                cursor.execute("""
                    UPDATE CONVITES_CONTATO
                    SET STATUS_ENVIO = 'ERRO',
                        ERRO_ENVIO = :1
                    WHERE TOKEN = :2
                """, (str(erro)[:500], token))
                conexao.commit()
            except Exception:
                pass

        flash(f"Erro ao enviar e-mail: {erro}", "danger")
        print("ERRO SMTP:", erro)
        return redirect(url_for("listar"))

    finally:
        cursor.close()
        conexao.close()

    return render_template("enviar_convite.html", cabo=cabo)

@app.route("/convite/<token>", methods=["GET", "POST"])
def cadastro_por_convite(token):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("login"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT ID, EMAIL, CABO_ID, STATUS, DATA_EXPIRACAO
            FROM CONVITES_CONTATO
            WHERE TOKEN = :1
        """, (token,))
        convite = cursor.fetchone()

        if not convite:
            flash("Convite inválido.", "danger")
            return redirect(url_for("login"))

        convite_id = convite[0]
        email = convite[1]
        cabo_id = convite[2]
        status = convite[3]
        data_expiracao = convite[4]

        if data_expiracao and data_expiracao < datetime.now():
            cursor.execute("""
                UPDATE CONVITES_CONTATO
                SET STATUS = 'EXPIRADO'
                WHERE ID = :1
                  AND STATUS = 'PENDENTE'
            """, (convite_id,))
            conexao.commit()
            flash("Este convite está expirado.", "warning")
            return redirect(url_for("login"))

        if status != "PENDENTE":
            flash("Este convite já foi utilizado ou não está mais disponível.", "warning")
            return redirect(url_for("login"))

        cursor.execute("""
            SELECT NOME
            FROM CABOS_ELEITORAIS
            WHERE ID = :1
        """, (cabo_id,))
        cabo = cursor.fetchone()

        if not cabo:
            flash("Liderança vinculada ao convite não encontrada.", "danger")
            return redirect(url_for("login"))

        nome_cabo = cabo[0]

        if request.method == "POST":
            nome = request.form.get("nome", "").strip()
            telefone = request.form.get("telefone", "").strip()
            endereco = request.form.get("endereco", "").strip()
            regiao = request.form.get("regiao_administrativa", "").strip()
            data_nascimento = converter_data(request.form.get("data_nascimento", "").strip())
            cep = request.form.get("cep", "").strip()
            consentiu = request.form.get("consentiu_contato", "N")
            observacao = request.form.get("observacao", "").strip()

            if not nome:
                flash("O nome é obrigatório.", "warning")
                return render_template("cadastro_por_convite.html", email=email, nome_cabo=nome_cabo)

            cursor.execute("""
                SELECT ID
                FROM CONTATOS_CAMPANHA
                WHERE UPPER(NOME) = UPPER(:1)
                  AND CABO_ID = :2
            """, (nome, cabo_id))
            registro_existente = cursor.fetchone()

            if registro_existente:
                flash("Já existe um contato cadastrado para esta liderança com esse nome.", "warning")
                return render_template("cadastro_por_convite.html", email=email, nome_cabo=nome_cabo)

            cursor.execute("""
                INSERT INTO CONTATOS_CAMPANHA
                (NOME, EMAIL, TELEFONE, ENDERECO, REGIAO_ADMINISTRATIVA, CABO_ID,
                CONSENTIU_CONTATO, OBSERVACAO, DATA_NASCIMENTO, CEP)
                VALUES (:1, :2, :3, :4, :5, :6, :7, :8, :9, :10)
            """, (
                nome,
                email,
                telefone,
                endereco,
                regiao,
                cabo_id,
                consentiu,
                observacao,
                data_nascimento,
                cep
            ))
            cursor.execute("""
                UPDATE CONVITES_CONTATO
                SET STATUS = 'USADO',
                    DATA_USO = SYSDATE
                WHERE ID = :1
            """, (convite_id,))

            conexao.commit()
            flash("Cadastro concluído com sucesso.", "success")
            return redirect(url_for("login"))

    except oracledb.Error as erro:
        flash(f"Erro ao processar cadastro por convite: {erro}", "danger")
        print("ERRO CADASTRO POR CONVITE:", erro)
        return redirect(url_for("login"))
    finally:
        cursor.close()
        conexao.close()

    return render_template("cadastro_por_convite.html", email=email, nome_cabo=nome_cabo)

@app.route("/listar-convites")
@login_required
def listar_convites():
    termo = request.args.get("busca", "").strip()
    status = request.args.get("status", "").strip()
    status_envio = request.args.get("status_envio", "").strip()
    cabo_id = request.args.get("cabo_id", "").strip()

    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return render_template(
            "listar_convites.html",
            registros=[],
            termo=termo,
            status=status,
            status_envio=status_envio,
            cabo_id=cabo_id,
            cabos=[]
        )

    cursor = conexao.cursor()

    try:
        if tipo_acesso == "CABO":
            cursor.execute("""
                SELECT ID, NOME
                FROM CABOS_ELEITORAIS
                WHERE ID = :1
                ORDER BY NOME
            """, (cabo_sessao,))
        else:
            cursor.execute("""
                SELECT ID, NOME
                FROM CABOS_ELEITORAIS
                ORDER BY NOME
            """)
        cabos = cursor.fetchall()

        sql = """
            SELECT
                cc.ID,
                cc.EMAIL,
                c.NOME,
                cc.STATUS,
                cc.STATUS_ENVIO,
                cc.DATA_ENVIO,
                cc.DATA_EXPIRACAO,
                cc.DATA_USO,
                cc.ERRO_ENVIO
            FROM CONVITES_CONTATO cc
            JOIN CABOS_ELEITORAIS c ON c.ID = cc.CABO_ID
            WHERE 1=1
        """
        params = {}

        if termo:
            sql += " AND UPPER(cc.EMAIL) LIKE :busca"
            params["busca"] = f"%{termo.upper()}%"

        if status:
            sql += " AND cc.STATUS = :status"
            params["status"] = status

        if status_envio:
            sql += " AND cc.STATUS_ENVIO = :status_envio"
            params["status_envio"] = status_envio

        if cabo_id:
            sql += " AND cc.CABO_ID = :cabo_id"
            params["cabo_id"] = int(cabo_id)

        if tipo_acesso == "CABO":
            sql += " AND cc.CABO_ID = :cabo_sessao"
            params["cabo_sessao"] = cabo_sessao

        sql += " ORDER BY cc.ID DESC"

        cursor.execute(sql, params)
        registros = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao listar convites: {erro}", "danger")
        print("ERRO AO LISTAR CONVITES:", erro)
        registros = []
        cabos = []
    finally:
        cursor.close()
        conexao.close()

    return render_template(
        "listar_convites.html",
        registros=registros,
        termo=termo,
        status=status,
        status_envio=status_envio,
        cabo_id=cabo_id,
        cabos=cabos
    )    
    
@app.route("/reenviar-convite/<int:convite_id>")
@login_required
def reenviar_convite(convite_id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_convites"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT cc.EMAIL, cc.TOKEN, c.NOME
            FROM CONVITES_CONTATO cc
            JOIN CABOS_ELEITORAIS c ON c.ID = cc.CABO_ID
            WHERE cc.ID = :1
        """, (convite_id,))
        convite = cursor.fetchone()

        if not convite:
            flash("Convite não encontrado.", "warning")
            return redirect(url_for("listar_convites"))

        email = convite[0]
        token = convite[1]
        nome_cabo = convite[2]

        link_convite = f"{BASE_URL}/convite/{token}"
        recusados = enviar_email_convite(email, nome_cabo, link_convite)

        if recusados:
            cursor.execute("""
                UPDATE CONVITES_CONTATO
                SET STATUS_ENVIO = 'RECUSADO',
                    ERRO_ENVIO = :1
                WHERE ID = :2
            """, (str(recusados)[:500], convite_id))
            conexao.commit()
            flash("O servidor recusou o destinatário.", "warning")
        else:
            cursor.execute("""
                UPDATE CONVITES_CONTATO
                SET STATUS_ENVIO = 'ENVIADO',
                    ERRO_ENVIO = NULL
                WHERE ID = :1
            """, (convite_id,))
            conexao.commit()
            flash("Convite reenviado com sucesso.", "success")

    except Exception as erro:
        try:
            cursor.execute("""
                UPDATE CONVITES_CONTATO
                SET STATUS_ENVIO = 'ERRO',
                    ERRO_ENVIO = :1
                WHERE ID = :2
            """, (str(erro)[:500], convite_id))
            conexao.commit()
        except Exception:
            pass

        flash(f"Erro ao reenviar convite: {erro}", "danger")
        print("ERRO REENVIAR CONVITE:", erro)
    finally:
        cursor.close()
        conexao.close()

    return redirect(url_for("listar_convites"))

@app.route("/cancelar-convite/<int:convite_id>")
@login_required
@perfil_required("ADMIN", "CHEFE_GABINETE", "SECRETARIA")
def cancelar_convite(convite_id):
    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_convites"))

    cursor = conexao.cursor()

    try:
        cursor.execute("""
            SELECT ID, STATUS
            FROM CONVITES_CONTATO
            WHERE ID = :1
        """, (convite_id,))
        convite = cursor.fetchone()

        if not convite:
            flash("Convite não encontrado.", "warning")
            return redirect(url_for("listar_convites"))

        status_atual = convite[1]

        if status_atual == "USADO":
            flash("Não é possível cancelar um convite que já foi utilizado.", "warning")
            return redirect(url_for("listar_convites"))

        if status_atual == "CANCELADO":
            flash("Este convite já está cancelado.", "info")
            return redirect(url_for("listar_convites"))

        cursor.execute("""
            UPDATE CONVITES_CONTATO
            SET STATUS = 'CANCELADO',
                ERRO_ENVIO = 'Convite cancelado pelo usuário'
            WHERE ID = :1
        """, (convite_id,))

        conexao.commit()

        flash("Convite cancelado com sucesso.", "success")

    except oracledb.Error as erro:
        conexao.rollback()
        flash(f"Erro ao cancelar convite: {erro}", "danger")
        print("ERRO AO CANCELAR CONVITE:", erro)

    finally:
        cursor.close()
        conexao.close()

    return redirect(url_for("listar_convites"))

@app.route("/exportar-convites-excel")
@login_required
def exportar_convites_excel():
    termo = request.args.get("busca", "").strip()
    status = request.args.get("status", "").strip()
    status_envio = request.args.get("status_envio", "").strip()
    cabo_id = request.args.get("cabo_id", "").strip()

    tipo_acesso = session.get("tipo_acesso")
    cabo_sessao = session.get("cabo_id")

    conexao = conectar_oracle()
    if conexao is None:
        flash("Não foi possível conectar ao banco de dados.", "danger")
        return redirect(url_for("listar_convites"))

    cursor = conexao.cursor()

    try:
        sql = """
            SELECT
                cc.ID,
                cc.EMAIL,
                c.NOME AS LIDERANCA,
                cc.STATUS,
                cc.STATUS_ENVIO,
                cc.DATA_ENVIO,
                cc.DATA_EXPIRACAO,
                cc.DATA_USO,
                cc.ERRO_ENVIO
            FROM CONVITES_CONTATO cc
            JOIN CABOS_ELEITORAIS c ON c.ID = cc.CABO_ID
            WHERE 1=1
        """
        params = {}

        if termo:
            sql += " AND UPPER(cc.EMAIL) LIKE :busca"
            params["busca"] = f"%{termo.upper()}%"

        if status:
            sql += " AND cc.STATUS = :status"
            params["status"] = status

        if status_envio:
            sql += " AND cc.STATUS_ENVIO = :status_envio"
            params["status_envio"] = status_envio

        if cabo_id:
            sql += " AND cc.CABO_ID = :cabo_id"
            params["cabo_id"] = int(cabo_id)

        if tipo_acesso == "CABO":
            sql += " AND cc.CABO_ID = :cabo_sessao"
            params["cabo_sessao"] = cabo_sessao

        sql += " ORDER BY cc.ID DESC"

        cursor.execute(sql, params)
        registros = cursor.fetchall()

    except oracledb.Error as erro:
        flash(f"Erro ao exportar convites: {erro}", "danger")
        print("ERRO AO EXPORTAR CONVITES:", erro)
        return redirect(url_for("listar_convites"))
    finally:
        cursor.close()
        conexao.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Convites Enviados"

    sheet.append([
        "ID",
        "E-mail",
        "Liderança",
        "Status do Convite",
        "Status do Envio",
        "Data de Envio",
        "Data de Expiração",
        "Data de Uso",
        "Erro de Envio"
    ])

    for registro in registros:
        linha = list(registro)

        if linha[5]:
            linha[5] = linha[5].strftime("%d/%m/%Y %H:%M")

        if linha[6]:
            linha[6] = linha[6].strftime("%d/%m/%Y %H:%M")

        if linha[7]:
            linha[7] = linha[7].strftime("%d/%m/%Y %H:%M")

        sheet.append(linha)

    arquivo_excel = BytesIO()
    workbook.save(arquivo_excel)
    arquivo_excel.seek(0)

    return send_file(
        arquivo_excel,
        as_attachment=True,
        download_name="convites_enviados.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
if __name__ == "__main__":
    debug = os.getenv("FLASK_DEBUG", "False").lower() == "true"
    app.run(debug=debug)
