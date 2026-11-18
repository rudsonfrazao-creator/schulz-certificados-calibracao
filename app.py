# -*- coding: utf-8 -*-
"""
Schulz S.A. - Automacao de Certificados de Calibracao (MVP - AV 05)
Projeto Aplicado III - UniSENAI

Modulos:
  1. Ingestao de dados  (leitura e higienizacao do .xlsx)
  2. Motor metrologico  (media, erro sistematico, comparacao com o LEA)
  3. Emissao            (classificacao, certificado PDF, QR Code, validacao mobile)
"""
import io
import json
import os
import re
import threading
import unicodedata
import uuid
from datetime import datetime

from flask import Flask, abort, redirect, render_template_string, request, send_file, url_for
from openpyxl import load_workbook
from reportlab.graphics import renderPDF
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
CERT_DIR = os.path.join(DATA_DIR, "certificados")
DB_FILE = os.path.join(DATA_DIR, "registros.json")
os.makedirs(CERT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB
_lock = threading.Lock()

MIN_LEITURAS = 3  # minimo de leituras validas para emitir certificado


# ---------------------------------------------------------------------------
# 1. MODULO DE INGESTAO DE DADOS
# ---------------------------------------------------------------------------
def normalizar_cabecalho(texto):
    """'Valor de Referência (mm)' -> 'valor_de_referencia_mm'"""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"[^a-zA-Z0-9]+", "_", t).strip("_").lower()
    t = re.sub(r"_(de|do|da)_", "_", t)   # 'valor_de_referencia' -> 'valor_referencia'
    return t


def higienizar_texto(valor):
    """Remove caracteres especiais/controle de campos de texto; nulo -> ''."""
    if valor is None:
        return ""
    t = unicodedata.normalize("NFKD", str(valor))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"[^\w\s\-\.\/]", "", t)
    return re.sub(r"\s+", " ", t).strip()


def converter_numero(valor):
    """Converte celula em float. Retorna None para nulo/invalido.
    Aceita virgula decimal e remove ruido (ex.: '10,02 mm' -> 10.02)."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    t = str(valor).strip().replace("\u00a0", "")
    if t == "" or t.lower() in {"nan", "na", "n/a", "null", "none", "-", "--"}:
        return None
    t = re.sub(r"[^\d,\.\-eE+]", "", t)
    if "," in t and "." in t:          # 1.234,56 -> 1234.56
        t = t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def ler_planilha(arquivo):
    """Le o .xlsx e devolve (instrumentos, avisos).
    Colunas esperadas (ordem livre): instrumento, identificacao, unidade,
    valor_referencia, lea, leitura_1 ... leitura_N"""
    wb = load_workbook(arquivo, data_only=True, read_only=True)
    ws = wb.active
    linhas = list(ws.iter_rows(values_only=True))
    if len(linhas) < 2:
        raise ValueError("A planilha esta vazia ou sem linhas de dados.")

    cab = [normalizar_cabecalho(c) for c in linhas[0]]
    idx_leituras = [i for i, c in enumerate(cab) if c.startswith("leitura")]
    obrig = ["instrumento", "valor_referencia", "lea"]
    faltando = [c for c in obrig if c not in cab]
    if faltando or not idx_leituras:
        raise ValueError(
            "Colunas obrigatorias ausentes: "
            + ", ".join(faltando + ([] if idx_leituras else ["leitura_1..N"]))
        )

    def col(nome):
        return cab.index(nome) if nome in cab else None

    avisos, instrumentos = [], []
    for n, linha in enumerate(linhas[1:], start=2):
        if all(c is None or str(c).strip() == "" for c in linha):
            continue  # linha totalmente vazia
        ref = converter_numero(linha[col("valor_referencia")])
        lea = converter_numero(linha[col("lea")])
        leituras_brutas = [linha[i] if i < len(linha) else None for i in idx_leituras]
        leituras = [x for x in (converter_numero(v) for v in leituras_brutas) if x is not None]
        descartadas = len(leituras_brutas) - len(leituras)
        if descartadas:
            avisos.append(f"Linha {n}: {descartadas} leitura(s) nula(s)/invalida(s) descartada(s).")

        ci = col("identificacao")
        cu = col("unidade")
        instrumentos.append({
            "linha": n,
            "instrumento": higienizar_texto(linha[col("instrumento")]) or f"Instrumento linha {n}",
            "identificacao": higienizar_texto(linha[ci]) if ci is not None else f"ID-{n:03d}",
            "unidade": higienizar_texto(linha[cu]) if cu is not None else "",
            "referencia": ref,
            "lea": lea,
            "leituras": leituras,
        })
    wb.close()
    return instrumentos, avisos


# ---------------------------------------------------------------------------
# 2. MOTOR DE CALCULOS METROLOGICOS
# ---------------------------------------------------------------------------
def calcular(inst):
    """media = soma(leituras)/n ; erro_sistematico = media - referencia ;
    APROVADO se |erro_sistematico| <= LEA."""
    leit, ref, lea = inst["leituras"], inst["referencia"], inst["lea"]
    if ref is None or lea is None or lea < 0:
        inst.update(media=None, erro=None, status="INVALIDO",
                    motivo="Valor de referencia ou LEA ausente/invalido.")
        return inst
    if len(leit) < MIN_LEITURAS:
        inst.update(media=None, erro=None, status="INVALIDO",
                    motivo=f"Menos de {MIN_LEITURAS} leituras validas.")
        return inst
    media = sum(leit) / len(leit)
    erro = media - ref
    aprovado = abs(erro) <= lea
    inst.update(media=media, erro=erro, n=len(leit),
                status="APROVADO" if aprovado else "REPROVADO", motivo="")
    return inst


# ---------------------------------------------------------------------------
# 3. EMISSAO E RASTREABILIDADE
# ---------------------------------------------------------------------------
def carregar_db():
    if not os.path.exists(DB_FILE):
        return {}
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar_registro(reg):
    with _lock:
        db = carregar_db()
        db[reg["codigo"]] = reg
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)


def base_url():
    return os.environ.get("BASE_URL", request.host_url).rstrip("/")


def fmt(v, casas=4):
    return "-" if v is None else f"{v:.{casas}f}".replace(".", ",")


def gerar_pdf(reg, url_validacao):
    caminho = os.path.join(CERT_DIR, f"{reg['codigo']}.pdf")
    c = canvas.Canvas(caminho, pagesize=A4)
    w, h = A4
    azul = colors.HexColor("#0B3D91")
    cor_status = colors.HexColor("#1B7F3B") if reg["status"] == "APROVADO" else colors.HexColor("#B3261E")

    c.setFillColor(azul)
    c.rect(0, h - 30 * mm, w, 30 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(20 * mm, h - 15 * mm, "CERTIFICADO DE CALIBRACAO")
    c.setFont("Helvetica", 10)
    c.drawString(20 * mm, h - 22 * mm, "Schulz S.A. - Laboratorio de Metrologia (MVP academico)")

    c.setFillColor(colors.black)
    y = h - 45 * mm
    linhas = [
        ("Codigo do certificado", reg["codigo"]),
        ("Instrumento", reg["instrumento"]),
        ("Identificacao", reg["identificacao"]),
        ("Data de emissao", reg["emitido_em"]),
        ("Valor de referencia", f"{fmt(reg['referencia'])} {reg['unidade']}"),
        ("Numero de leituras validas", str(reg.get("n", len(reg["leituras"])))),
        ("Media das leituras", f"{fmt(reg['media'])} {reg['unidade']}"),
        ("Erro sistematico", f"{fmt(reg['erro'])} {reg['unidade']}"),
        ("Limite de Erro Admissivel (LEA)", f"+/- {fmt(reg['lea'])} {reg['unidade']}"),
    ]
    for rot, val in linhas:
        c.setFont("Helvetica-Bold", 10)
        c.drawString(20 * mm, y, rot + ":")
        c.setFont("Helvetica", 10)
        c.drawString(88 * mm, y, str(val))
        y -= 8 * mm

    y -= 4 * mm
    c.setFillColor(cor_status)
    c.roundRect(20 * mm, y - 6 * mm, 60 * mm, 14 * mm, 3 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(50 * mm, y - 1.5 * mm, reg["status"])

    # QR Code de validacao
    qr = QrCodeWidget(url_validacao)
    x1, y1, x2, y2 = qr.getBounds()
    tam = 40 * mm
    d = Drawing(tam, tam, transform=[tam / (x2 - x1), 0, 0, tam / (y2 - y1), 0, 0])
    d.add(qr)
    renderPDF.draw(d, c, w - 65 * mm, 40 * mm)
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8)
    c.drawCentredString(w - 45 * mm, 35 * mm, "Escaneie para validar")
    c.drawString(20 * mm, 30 * mm, "Validacao: " + url_validacao)
    c.setFont("Helvetica-Oblique", 7)
    c.drawString(20 * mm, 20 * mm, "Criterio: APROVADO se |erro sistematico| <= LEA. Documento gerado automaticamente.")
    c.save()
    return caminho


def emitir(inst):
    codigo = uuid.uuid4().hex[:12].upper()
    reg = dict(inst, codigo=codigo, emitido_em=datetime.now().strftime("%d/%m/%Y %H:%M"))
    if reg["status"] in ("APROVADO", "REPROVADO"):
        gerar_pdf(reg, f"{base_url()}/validar/{codigo}")
    salvar_registro(reg)
    return reg


# ---------------------------------------------------------------------------
# INTERFACE WEB
# ---------------------------------------------------------------------------
CSS = """<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:900px;margin:30px auto;padding:0 16px;color:#222}
h1{color:#0B3D91}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px 8px;font-size:14px;text-align:left}
th{background:#0B3D91;color:#fff}.ok{color:#1B7F3B;font-weight:bold}.no{color:#B3261E;font-weight:bold}
.card{border:1px solid #ddd;border-radius:10px;padding:20px;margin:16px 0}.btn{background:#0B3D91;color:#fff;border:0;padding:10px 18px;border-radius:6px;cursor:pointer;text-decoration:none}
.aviso{background:#FFF4CE;padding:8px 12px;border-radius:6px;margin:6px 0;font-size:13px}</style>
<meta name="viewport" content="width=device-width,initial-scale=1">"""

HOME = CSS + """<title>Schulz - Certificados</title><h1>Schulz S.A. - Automacao de Certificados</h1>
<div class="card"><p>Envie a planilha <b>.xlsx</b> com as leituras. Colunas: <i>instrumento, identificacao, unidade,
valor_referencia, lea, leitura_1 ... leitura_N</i>.</p>
{% if erro %}<p class="no">{{ erro }}</p>{% endif %}
<form method="post" action="/processar" enctype="multipart/form-data">
<input type="file" name="planilha" accept=".xlsx" required> <button class="btn">Processar</button></form>
<p><a href="/modelo">Baixar planilha modelo</a></p></div>"""

RESULTADO = CSS + """<title>Resultado</title><h1>Resultado do processamento</h1>
{% for a in avisos %}<div class="aviso">{{ a }}</div>{% endfor %}
<table><tr><th>Instrumento</th><th>ID</th><th>Media</th><th>Erro sist.</th><th>LEA</th><th>Status</th><th>Certificado</th></tr>
{% for r in regs %}<tr><td>{{ r.instrumento }}</td><td>{{ r.identificacao }}</td><td>{{ f(r.media) }}</td><td>{{ f(r.erro) }}</td>
<td>{{ f(r.lea) }}</td><td class="{{ 'ok' if r.status=='APROVADO' else 'no' }}">{{ r.status }}{% if r.motivo %}<br><small>{{ r.motivo }}</small>{% endif %}</td>
<td>{% if r.status in ['APROVADO','REPROVADO'] %}<a href="/certificado/{{ r.codigo }}">PDF</a> | <a href="/validar/{{ r.codigo }}">Validar</a>{% else %}-{% endif %}</td></tr>{% endfor %}</table>
<p><a class="btn" href="/">Novo processamento</a></p>"""

VALIDAR = CSS + """<title>Validacao</title><h1>Validacao de Certificado</h1><div class="card">
{% if r %}<p>Certificado <b>{{ r.codigo }}</b> <span class="ok">AUTENTICO</span></p>
<p>Instrumento: {{ r.instrumento }} ({{ r.identificacao }})<br>Emitido em: {{ r.emitido_em }}<br>
Media: {{ f(r.media) }} {{ r.unidade }} | Erro: {{ f(r.erro) }} | LEA: +/-{{ f(r.lea) }}</p>
<h2 class="{{ 'ok' if r.status=='APROVADO' else 'no' }}">{{ r.status }}</h2>
{% else %}<p class="no">Certificado NAO encontrado. Documento nao reconhecido.</p>{% endif %}</div>"""


@app.route("/")
def home():
    return render_template_string(HOME, erro=request.args.get("erro"))


@app.route("/processar", methods=["POST"])
def processar():
    arq = request.files.get("planilha")
    if not arq or not arq.filename.lower().endswith(".xlsx"):
        return redirect(url_for("home", erro="Envie um arquivo .xlsx valido."))
    try:
        instrumentos, avisos = ler_planilha(arq)
    except Exception as e:  # planilha corrompida / colunas faltando
        return redirect(url_for("home", erro=str(e)))
    regs = [emitir(calcular(i)) for i in instrumentos]
    return render_template_string(RESULTADO, regs=regs, avisos=avisos, f=fmt)


@app.route("/certificado/<codigo>")
def certificado(codigo):
    if not re.fullmatch(r"[A-F0-9]{12}", codigo):
        abort(404)
    caminho = os.path.join(CERT_DIR, f"{codigo}.pdf")
    if not os.path.exists(caminho):
        abort(404)
    return send_file(caminho, mimetype="application/pdf")


@app.route("/validar/<codigo>")
def validar(codigo):
    return render_template_string(VALIDAR, r=carregar_db().get(codigo.upper()), f=fmt)


@app.route("/modelo")
def modelo():
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["Instrumento", "Identificacao", "Unidade", "Valor de Referencia", "LEA",
               "Leitura 1", "Leitura 2", "Leitura 3", "Leitura 4", "Leitura 5"])
    ws.append(["Paquimetro", "PQ-001", "mm", 100.000, 0.030, 100.01, 100.02, 99.99, 100.00, 100.01])
    ws.append(["Micrometro", "MC-014", "mm", 25.000, 0.004, "25,008", 25.007, None, 25.009, "25,008 mm"])
    ws.append(["Manometro", "MN-203", "bar", 10.000, 0.050, 10.09, 10.11, 10.10, 10.12, 10.08])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="planilha_modelo.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
