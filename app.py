import base64
import hmac
import html
import io
import json
import os
import tempfile
import traceback
import uuid
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from fpdf import FPDF
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw
from streamlit_drawable_canvas import st_canvas

st.set_page_config(page_title="Controle de Equipamentos", layout="wide")

FUSO = ZoneInfo("America/Sao_Paulo")
FORMATO = "%Y-%m-%d %H:%M:%S"

# Ordem EXATA das colunas da aba "Registros" (A até K)
# A coluna K (Transferencias) é criada sozinha na primeira transferência.
COLUNAS = [
    "ID", "Equipamento", "Lacre", "Colaborador",
    "Data_Hora_Saida", "Assinatura_Saida",
    "Data_Hora_Entrada", "Assinatura_Entrada",
    "Status", "Tempo_Uso_Horas",
    "Transferencias",
]

# Cole o link da planilha aqui OU coloque em Secrets: planilha_url = "https://..."
PLANILHA_URL = st.secrets.get("https://docs.google.com/spreadsheets/d/1i5uOxp-hNCAvvbBX1580AKGpp9zWYu2DHBcK43272Sw/edit?gid=483892025#gid=483892025")

# Limite de uma célula do Google Sheets = 50.000 caracteres
LIMITE_CELULA = 45000

# ---- Regras de autorização (senha do responsável) ----
# A senha fica em Secrets:  senha_responsavel = "sua-senha-aqui"
PRAZO_DEVOLUCAO_HORAS = 12      # passou disso sem devolver -> nova retirada precisa de autorização
VALIDADE_LIBERACAO_MIN = 15     # quanto tempo a liberação da assinatura de devolução vale
ABA_AUTORIZACOES = "Autorizacoes"
COLUNAS_AUT = [
    "Data_Hora", "Tipo", "ID_Registro", "Colaborador",
    "Equipamento", "Lacre", "Justificativa", "Pendencias",
]
TIPO_RETIRADA = "Retirada com devolução atrasada"
TIPO_LIB_DEVOLUCAO = "Liberação de assinatura (devolução)"
TIPO_TRANSFERENCIA = "Transferência de equipamento"

# Chaves para limpar o quadro de assinatura depois de salvar
for chave in ("saida_key", "entrada_key", "aut_key", "lib_key", "transf_key"):
    if chave not in st.session_state:
        st.session_state[chave] = 0


# ---------------------------------------------------------------- Google Sheets
@st.cache_resource
def conectar_planilha():
    credenciais = dict(st.secrets["gcp_service_account"])
    escopo = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(credenciais, scopes=escopo)
    client = gspread.authorize(creds)
    return client.open_by_url(PLANILHA_URL)


def aba_registros():
    return conectar_planilha().worksheet("Registros")


def carregar_dados():
    """Lê tudo como TEXTO (evita o Sheets converter número/data/vírgula)."""
    valores = aba_registros().get_all_values()
    if len(valores) <= 1:
        return pd.DataFrame(columns=COLUNAS)
    linhas = []
    for r in valores[1:]:
        r = (r + [""] * len(COLUNAS))[: len(COLUNAS)]
        if any(c.strip() for c in r):
            linhas.append(r)
    return pd.DataFrame(linhas, columns=COLUNAS)


@st.cache_data(ttl=120, show_spinner=False)
def ler_lista(nome_aba, colunas_preferidas, qtd_colunas):
    """Lê as abas Ferramentas / Colaboradores. Se não achar, devolve lista vazia."""
    try:
        valores = conectar_planilha().worksheet(nome_aba).get_all_values()
    except Exception:
        return []
    if len(valores) <= 1:
        return []
    cab = [c.strip().lower() for c in valores[0]]
    idx = []
    for nome in colunas_preferidas:
        idx.append(cab.index(nome.lower()) if nome.lower() in cab else None)
    if any(i is None for i in idx):
        idx = list(range(qtd_colunas))
    saida = []
    for r in valores[1:]:
        r = r + [""] * (max(idx) + 1 - len(r)) if len(r) <= max(idx) else r
        item = tuple(r[i].strip() for i in idx)
        if item[0]:
            saida.append(item)
    return saida


# ---------------------------------------------------------------- Assinatura
LARGURA_ASS, ALTURA_ASS = 300, 120
ESCALA = 3  # desenha em resolução maior e reduz depois (traço mais suave)


def _json_do_canvas(resultado):
    """Pega o json_data do canvas, aceitando dict ou texto JSON."""
    dados = getattr(resultado, "json_data", None)
    if isinstance(dados, (str, bytes)):
        try:
            dados = json.loads(dados)
        except ValueError:
            return {}
    return dados if isinstance(dados, dict) else {}


def _pontos_do_comando(comandos):
    """Converte comandos de caminho (M, L, Q, C) em uma lista de pontos."""
    pontos, atual = [], None
    for c in comandos:
        if not isinstance(c, (list, tuple)) or not c:
            continue
        tipo = str(c[0]).upper()
        try:
            v = [float(n) for n in c[1:]]
        except (TypeError, ValueError):
            continue
        if tipo in ("M", "L") and len(v) >= 2:
            atual = (v[0], v[1])
            pontos.append(atual)
        elif tipo == "Q" and len(v) >= 4 and atual is not None:
            c1, fim = (v[0], v[1]), (v[2], v[3])
            for i in range(1, 9):
                t = i / 8
                pontos.append((
                    (1 - t) ** 2 * atual[0] + 2 * (1 - t) * t * c1[0] + t ** 2 * fim[0],
                    (1 - t) ** 2 * atual[1] + 2 * (1 - t) * t * c1[1] + t ** 2 * fim[1],
                ))
            atual = fim
        elif tipo == "C" and len(v) >= 6 and atual is not None:
            c1, c2, fim = (v[0], v[1]), (v[2], v[3]), (v[4], v[5])
            for i in range(1, 9):
                t = i / 8
                a = (1 - t) ** 3
                b = 3 * (1 - t) ** 2 * t
                c_ = 3 * (1 - t) * t ** 2
                d = t ** 3
                pontos.append((
                    a * atual[0] + b * c1[0] + c_ * c2[0] + d * fim[0],
                    a * atual[1] + b * c1[1] + c_ * c2[1] + d * fim[1],
                ))
            atual = fim
    return pontos


def _tracos_do_canvas(resultado):
    """Lista de traços (cada um é uma lista de pontos). Funciona mesmo que a
    biblioteca use 'path' ou 'Path' como tipo do objeto."""
    objetos = _json_do_canvas(resultado).get("objects") or []
    tracos = []
    for o in objetos:
        if isinstance(o, dict) and o.get("path"):
            pts = _pontos_do_comando(o["path"])
            if pts:
                tracos.append(pts)
    return tracos


def _tracos_da_imagem(resultado):
    """Plano B: se a biblioteca ainda entregar a imagem, usa os pixels escuros."""
    try:
        dados = resultado.image_data
    except Exception:
        return None
    if dados is None:
        return None
    img = Image.fromarray(dados.astype("uint8"), "RGBA")
    fundo = Image.new("RGB", img.size, "white")
    fundo.paste(img, mask=img.split()[3])
    cinza = fundo.convert("L")
    if cinza.getextrema()[0] > 200:
        return None
    return cinza


def _codifica(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    texto = base64.b64encode(buf.getvalue()).decode("ascii")
    return texto if len(texto) <= LIMITE_CELULA else None


def assinatura_para_base64(resultado):
    """Redesenha a assinatura em PNG pequeno (base64). None se o quadro estiver vazio.
    O traço é centralizado e ajustado ao tamanho do quadro."""
    if resultado is None:
        return None

    tracos = _tracos_do_canvas(resultado)
    if tracos:
        xs = [p[0] for t in tracos for p in t]
        ys = [p[1] for t in tracos for p in t]
        minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
        margem = 8
        larg = max(maxx - minx, 1.0)
        alt = max(maxy - miny, 1.0)
        fator = min((LARGURA_ASS - 2 * margem) / larg, (ALTURA_ASS - 2 * margem) / alt, 3.0)
        off_x = (LARGURA_ASS - larg * fator) / 2 - minx * fator
        off_y = (ALTURA_ASS - alt * fator) / 2 - miny * fator

        img = Image.new("L", (LARGURA_ASS * ESCALA, ALTURA_ASS * ESCALA), 255)
        draw = ImageDraw.Draw(img)
        espessura = 2 * ESCALA
        for t in tracos:
            pts = [((x * fator + off_x) * ESCALA, (y * fator + off_y) * ESCALA) for x, y in t]
            if len(pts) == 1:
                x, y = pts[0]
                draw.ellipse([x - espessura, y - espessura, x + espessura, y + espessura], fill=0)
            else:
                draw.line(pts, fill=0, width=espessura, joint="curve")
        img = img.resize((LARGURA_ASS, ALTURA_ASS), Image.LANCZOS)
        return _codifica(img)

    cinza = _tracos_da_imagem(resultado)
    if cinza is not None:
        return _codifica(cinza.resize((LARGURA_ASS, ALTURA_ASS), Image.LANCZOS))
    return None


def aviso_sem_assinatura(resultado=None):
    st.warning("Assine no quadro antes de registrar.")


def campo_assinatura(chave):
    return st_canvas(
        fill_color="rgba(255,255,255,0)",
        stroke_width=2,
        stroke_color="#000000",
        background_color="#f4f6f9",
        height=120,
        width=340,
        drawing_mode="freedraw",
        update_streamlit=True,
        key=chave,
    )


# ---------------------------------------------------------------- Utilidades
def agora():
    return datetime.now(FUSO).replace(tzinfo=None)


def para_float(x):
    try:
        return float(str(x).replace(",", "."))
    except ValueError:
        return 0.0


def formata_tempo(horas):
    total_min = int(round(para_float(horas) * 60))
    return f"{total_min // 60}h {total_min % 60:02d}min"


def chave_equip(equip, lacre):
    return f"{equip}|{lacre}"


# ---------------------------------------------------------------- PDF
def limpa(txt):
    return str(txt).encode("latin-1", "replace").decode("latin-1")


def data_br(texto):
    """2026-11-09 08:23:10 -> 09/11/2026 08:23"""
    try:
        return datetime.strptime(texto, FORMATO).strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return texto or ""


class RelatorioPDF(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(120, 120, 120)
        self.cell(95, 5, limpa(f"Gerado em {agora():%d/%m/%Y %H:%M}"), align="L")
        self.cell(0, 5, limpa("Página ") + f"{self.page_no()}/{{nb}}", align="R")
        self.set_text_color(0, 0, 0)


def cabe(pdf, texto, largura):
    """Corta o texto com '...' se ele não couber na largura (mm) da fonte atual."""
    texto = limpa(texto)
    if pdf.get_string_width(texto) <= largura:
        return texto
    while texto and pdf.get_string_width(texto + "...") > largura:
        texto = texto[:-1]
    return texto + "..."


def texto_em(pdf, x, y, largura, texto, fonte=("Helvetica", "", 10), cor=(20, 20, 20), align="L"):
    pdf.set_font(*fonte)
    pdf.set_text_color(*cor)
    pdf.set_xy(x, y)
    pdf.cell(largura, 5, cabe(pdf, texto, largura), align=align)


def desenha_assinatura(pdf, b64, x, y, larg, alt, texto_vazio):
    pdf.set_draw_color(190, 195, 205)
    pdf.set_line_width(0.3)
    pdf.rect(x, y, larg, alt)
    if not b64:
        texto_em(pdf, x, y + alt / 2 - 2.5, larg, texto_vazio,
                 fonte=("Helvetica", "I", 8), cor=(150, 150, 150), align="C")
        return
    caminho = None
    try:
        img = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGB")
        # Salva em arquivo temporário (método que já funcionou com o fpdf2 deste app)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            caminho = tmp.name
        img.save(caminho, format="PNG")
        h_img = alt - 4
        w_img = h_img * img.width / img.height
        if w_img > larg - 4:
            w_img = larg - 4
            h_img = w_img * img.height / img.width
        pdf.image(caminho, x=x + (larg - w_img) / 2, y=y + (alt - h_img) / 2, w=w_img, h=h_img)
    except Exception:
        traceback.print_exc()  # o motivo aparece em "Gerenciar aplicativo" > logs
        texto_em(pdf, x, y + alt / 2 - 2.5, larg, "assinatura inválida",
                 fonte=("Helvetica", "I", 8), cor=(150, 150, 150), align="C")
    finally:
        if caminho and os.path.exists(caminho):
            os.remove(caminho)


def quebra_linhas(pdf, texto, largura, fonte, max_linhas=5):
    """Quebra o texto em linhas que cabem na largura (mm). Corta com '...' se passar do limite."""
    pdf.set_font(*fonte)
    texto = limpa(" ".join(str(texto).split()))
    linhas, atual = [], ""
    for palavra in texto.split(" "):
        teste = (atual + " " + palavra).strip()
        if pdf.get_string_width(teste) <= largura:
            atual = teste
        else:
            if atual:
                linhas.append(atual)
            atual = palavra
    if atual:
        linhas.append(atual)
    if len(linhas) > max_linhas:
        linhas = linhas[:max_linhas]
        linhas[-1] = cabe(pdf, linhas[-1] + "...", largura)
    return linhas


def gerar_pdf(df, titulo_periodo, motivos=None):
    ESQ, LARG, ALT = 10, 190, 62  # margem, largura útil e altura de cada bloco (mm)
    CINZA_RUIM = (110, 115, 125)
    motivos = motivos or {}
    F_TIT = ("Helvetica", "B", 8.5)
    F_TXT = ("Helvetica", "", 8.5)
    ALT_LINHA = 4.6

    pdf = RelatorioPDF(format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(False)
    pdf.set_margins(ESQ, 10, ESQ)
    pdf.add_page()

    # ----- cabeçalho
    texto_em(pdf, ESQ, 10, LARG, "Relatório de Movimentação de Equipamentos",
             fonte=("Helvetica", "B", 16), cor=(20, 20, 20))
    texto_em(pdf, ESQ, 19, LARG, f"Período: {titulo_periodo}",
             fonte=("Helvetica", "", 10), cor=CINZA_RUIM)

    # ----- resumo
    total = len(df)
    devolvidos = int((df["Status"] == "Devolvido").sum())
    em_uso = int((df["Status"] == "Em uso").sum())
    horas = df["Tempo_Uso_Horas"].apply(para_float).sum()
    pdf.set_fill_color(240, 242, 246)
    pdf.rect(ESQ, 27, LARG, 13, style="F")
    resumo = [("MOVIMENTAÇÕES", str(total)), ("DEVOLVIDOS", str(devolvidos)),
              ("EM USO", str(em_uso)), ("TEMPO TOTAL DE USO", formata_tempo(horas))]
    larg_col = LARG / len(resumo)
    for i, (rotulo, valor) in enumerate(resumo):
        x = ESQ + 3 + i * larg_col
        texto_em(pdf, x, 28, larg_col - 3, rotulo, fonte=("Helvetica", "", 7), cor=CINZA_RUIM)
        texto_em(pdf, x, 32.5, larg_col - 3, valor, fonte=("Helvetica", "B", 11), cor=(20, 20, 20))
    pdf.set_y(46)

    # ----- um bloco por movimentação
    for _, r in df.iterrows():
        transfs = ler_transferencias(r["Transferencias"])

        # notas do bloco: (texto, fonte, cor, recuo). Retirada autorizada + transferências, com motivo.
        notas = []
        aut = motivos.get("retirada", {}).get(r["ID"])
        if aut:
            notas.append(("RETIRADA AUTORIZADA PELO RESPONSÁVEL (devolução em atraso)", F_TIT, (190, 40, 40), 0))
            for t in quebra_linhas(pdf, f"Motivo: {aut[0] or 'não informado'}", LARG - 10, F_TXT, 5):
                notas.append((t, F_TXT, (60, 60, 60), 4))
            if aut[1]:
                for t in quebra_linhas(pdf, f"Pendências: {aut[1]}", LARG - 10, F_TXT, 3):
                    notas.append((t, F_TXT, (60, 60, 60), 4))
        motivos_transf = motivos.get("transf", {}).get(r["ID"], [])
        for i, (de, para, dt) in enumerate(transfs):
            quando = data_br(dt.strftime(FORMATO)) if dt else "?"
            notas.append((f"TRANSFERIDO: {de} -> {para} em {quando}", F_TIT, (180, 110, 0), 0))
            mot = motivos_transf[i] if i < len(motivos_transf) else ""
            if mot:
                for t in quebra_linhas(pdf, f"Motivo: {mot}", LARG - 10, F_TXT, 5):
                    notas.append((t, F_TXT, (60, 60, 60), 4))
        extra = ALT_LINHA * len(notas) + (1 if notas else 0)
        alt_bloco = ALT + extra
        if pdf.get_y() + alt_bloco > 278:
            pdf.add_page()
            pdf.set_y(12)
        y0 = pdf.get_y()
        devolvido = r["Status"] == "Devolvido"

        # moldura e faixa do topo
        pdf.set_draw_color(200, 205, 214)
        pdf.set_line_width(0.3)
        pdf.rect(ESQ, y0, LARG, alt_bloco)
        pdf.set_fill_color(240, 242, 246)
        pdf.rect(ESQ, y0, LARG, 10, style="DF")

        texto_em(pdf, ESQ + 3, y0 + 2.5, 85, r["Equipamento"], fonte=("Helvetica", "B", 11))
        texto_em(pdf, ESQ + 92, y0 + 2.5, 40, f"Lacre: {r['Lacre']}", fonte=("Helvetica", "", 10))
        cor_status = (22, 128, 60) if devolvido else (180, 110, 0)
        texto_em(pdf, ESQ + LARG - 55, y0 + 2.5, 52, str(r["Status"]).upper(),
                 fonte=("Helvetica", "B", 10), cor=cor_status, align="R")

        # dados em quatro colunas (rótulo pequeno em cima, valor embaixo)
        campos = [
            ("COLABORADOR ATUAL" if transfs else "COLABORADOR", r["Colaborador"]),
            ("SAÍDA", data_br(r["Data_Hora_Saida"])),
            ("ENTRADA", data_br(r["Data_Hora_Entrada"]) if devolvido else "Ainda não devolvido"),
            ("TEMPO DE USO", formata_tempo(r["Tempo_Uso_Horas"]) if devolvido else "-"),
        ]
        larg_campo = (LARG - 6) / 4
        for i, (rotulo, valor) in enumerate(campos):
            x = ESQ + 3 + i * larg_campo
            texto_em(pdf, x, y0 + 13, larg_campo - 2, rotulo, fonte=("Helvetica", "", 7), cor=CINZA_RUIM)
            texto_em(pdf, x, y0 + 17.5, larg_campo - 2, valor, fonte=("Helvetica", "B", 9.5))

        pdf.set_draw_color(225, 228, 234)
        pdf.line(ESQ + 3, y0 + 25.5, ESQ + LARG - 3, y0 + 25.5)

        # notas: retirada autorizada e transferências, com os motivos
        for i, (txt, fonte, cor, recuo) in enumerate(notas):
            texto_em(pdf, ESQ + 3 + recuo, y0 + 27 + i * ALT_LINHA, LARG - 6 - recuo, txt,
                     fonte=fonte, cor=cor)

        # assinaturas lado a lado
        larg_ass = 89
        x1 = ESQ + 3
        x2 = x1 + larg_ass + 4
        ya = y0 + extra  # desce as assinaturas quando há transferências
        rot_saida = "ASSINATURA DE SAÍDA" + (f" ({transfs[0][0]})" if transfs else "")
        rot_entrada = "ASSINATURA DE ENTRADA" + (f" ({r['Colaborador']})" if transfs else "")
        texto_em(pdf, x1, ya + 27, larg_ass, rot_saida.upper(), fonte=("Helvetica", "", 7), cor=CINZA_RUIM)
        texto_em(pdf, x2, ya + 27, larg_ass, rot_entrada.upper(), fonte=("Helvetica", "", 7), cor=CINZA_RUIM)
        desenha_assinatura(pdf, r["Assinatura_Saida"], x1, ya + 31.5, larg_ass, 26, "sem assinatura")
        desenha_assinatura(
            pdf, r["Assinatura_Entrada"] if devolvido else "", x2, ya + 31.5, larg_ass, 26,
            "aguardando devolução" if not devolvido else "sem assinatura",
        )

        pdf.set_y(y0 + alt_bloco + 5)

    saida = pdf.output(dest="S")
    if isinstance(saida, str):
        saida = saida.encode("latin-1")
    return bytes(saida)


# ---------------------------------------------------------------- Interface (visual)
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
.stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp li, .stApp a,
.stApp button, .stApp input, .stApp textarea { font-family: 'Inter', sans-serif; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { max-width: 1200px; padding-top: 2.2rem; padding-bottom: 3rem; }

/* cartões: fundo e borda translucidos, funcionam no tema claro e no escuro */
[class*="st-key-card"] {
    background: rgba(128,128,128,.07); border: 1px solid rgba(128,128,128,.28);
    border-radius: 14px; padding: 1.2rem 1.3rem;
}
[class*="st-key-card"] h3 { margin: 0 0 .6rem 0; font-size: 1.25rem; font-weight: 600; }

/* título */
.tit { font-size: 2rem; font-weight: 700; margin: 0; line-height: 1.2; }
.sub { opacity: .75; margin: .2rem 0 0 0; }

/* indicadores */
.kpi { display: flex; align-items: center; gap: .9rem; background: rgba(128,128,128,.07);
       border: 1px solid rgba(128,128,128,.28); border-radius: 14px; padding: 1.1rem 1.3rem; }
.kpi-ico { width: 52px; height: 52px; border-radius: 12px; display: flex; align-items: center; justify-content: center; }
.kpi-t { opacity: .8; font-size: .95rem; }
.kpi-v { font-size: 1.9rem; font-weight: 700; line-height: 1.1; }

/* tabela de devoluções */
.th { font-size: .8rem; font-weight: 600; background: rgba(128,128,128,.16); padding: .5rem .6rem;
      border-radius: 8px; white-space: nowrap; }
.mut { opacity: .65; font-size: .85rem; }
.sep { border: 0; border-top: 1px solid rgba(128,128,128,.22); margin: .35rem 0; }
.badge { display: inline-block; padding: .22rem .7rem; border-radius: 999px; font-size: .8rem;
         font-weight: 600; white-space: nowrap; }
.b-ok { background: rgba(34,197,94,.2); color: #22c55e; }
.b-warn { background: rgba(245,158,11,.2); color: #f59e0b; }

/* campos e botões */
div[data-baseweb="select"] > div, div[data-baseweb="input"] > div { border-radius: 9px !important; }
.stButton > button { border-radius: 10px; font-weight: 600; padding: .5rem .8rem; white-space: nowrap; }
.stButton > button[kind="primary"] { background: #e11d2e; border: 1px solid #e11d2e; color: #fff; }
.stButton > button[kind="primary"]:hover { background: #bf1625; border-color: #bf1625; color: #fff; }

/* aviso de devolução pendente */
.aviso { border: 2px solid #d9480f; border-radius: 12px; padding: 1rem 1.1rem;
         background: rgba(217,72,15,.08); margin-bottom: .8rem; }
.aviso-tit { font-weight: 700; font-size: 1.05rem; color: #d9480f; margin-bottom: .3rem; }
.aviso-item { background: rgba(128,128,128,.1); border: 1px solid rgba(128,128,128,.28);
              border-radius: 10px; padding: .6rem .8rem; margin-top: .5rem; font-size: .92rem; line-height: 1.55; }

/* quadro de assinatura (fundo claro para o traço preto aparecer nos dois temas) */
iframe[title*="st_canvas"] { border: 1.5px dashed rgba(128,128,128,.7); border-radius: 12px; background: #f4f6f9; }
.ass-legenda { text-align: center; opacity: .65; font-size: .85rem; margin-top: -.2rem; }

/* indicadores do relatório */
[data-testid="stMetric"] { background: rgba(128,128,128,.07); border: 1px solid rgba(128,128,128,.28);
                           border-radius: 14px; padding: 1rem 1.2rem; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

ICONE_OK = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12.5l3 3 5-6"/></svg>'
ICONE_USO = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 8h6M9 12h6M9 16h3"/></svg>'
ICONE_ALERTA = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18v.01"/></svg>'


def card_indicador(icone, fundo, cor, titulo, valor):
    st.markdown(
        f'<div class="kpi"><div class="kpi-ico" style="background:{fundo};color:{cor}">{icone}</div>'
        f'<div><div class="kpi-t">{titulo}</div><div class="kpi-v">{valor}</div></div></div>',
        unsafe_allow_html=True,
    )


def esc(x):
    return html.escape(str(x))


def horas_desde(texto):
    """Horas desde a saída e a data da saída (None se a data estiver inválida)."""
    try:
        saida = datetime.strptime(texto, FORMATO)
    except (ValueError, TypeError):
        return 0.0, None
    return max((agora() - saida).total_seconds() / 3600, 0.0), saida


# ---------------------------------------------------------------- Autorizações (senha do responsável)
def senha_correta(digitada):
    """Confere a senha do responsável (Secrets: senha_responsavel). Mostra o erro na tela."""
    correta = str(st.secrets.get("senha_responsavel", ""))
    if not correta:
        st.error("Senha do responsável não configurada. Defina 'senha_responsavel' em Secrets.")
        return False
    if not hmac.compare_digest(str(digitada).encode("utf-8"), correta.encode("utf-8")):
        st.error("Senha incorreta.")
        return False
    return True


def aba_autorizacoes():
    """Aba de auditoria. Se ainda não existir, é criada sozinha com o cabeçalho."""
    planilha = conectar_planilha()
    try:
        return planilha.worksheet(ABA_AUTORIZACOES)
    except gspread.WorksheetNotFound:
        ws = planilha.add_worksheet(title=ABA_AUTORIZACOES, rows=1000, cols=len(COLUNAS_AUT))
        ws.append_row(COLUNAS_AUT, value_input_option="RAW")
        return ws


def registrar_autorizacao(tipo, id_reg, colab, equip, lacre, justificativa="", pendencias=""):
    aba_autorizacoes().append_row(
        [agora().strftime(FORMATO), tipo, id_reg, colab, equip, str(lacre), justificativa, pendencias],
        value_input_option="RAW",
    )


def carregar_motivos():
    """Lê a aba Autorizacoes e devolve os motivos por registro, para o relatório:
    {"transf": {id: [motivo da 1ª transferência, 2ª, ...]},
     "retirada": {id: (motivo, pendências)}}  (retirada liberada com devolução atrasada)"""
    saida = {"transf": {}, "retirada": {}}
    try:
        valores = aba_autorizacoes().get_all_values()
    except Exception:
        traceback.print_exc()
        return saida
    for r in valores[1:]:
        r = (r + [""] * len(COLUNAS_AUT))[: len(COLUNAS_AUT)]
        tipo, id_reg, just, pend = r[1], r[2], r[6].strip(), r[7].strip()
        if tipo == TIPO_TRANSFERENCIA:
            saida["transf"].setdefault(id_reg, []).append(just)
        elif tipo == TIPO_RETIRADA:
            saida["retirada"][id_reg] = (just, pend)
    return saida


def liberacao_valida(id_reg):
    """True se o almoxarifado liberou a assinatura deste registro há menos de VALIDADE_LIBERACAO_MIN."""
    valores = aba_autorizacoes().get_all_values()
    for r in reversed(valores[1:]):
        r = (r + [""] * len(COLUNAS_AUT))[: len(COLUNAS_AUT)]
        if r[1] == TIPO_LIB_DEVOLUCAO and r[2] == id_reg:
            try:
                quando = datetime.strptime(r[0], FORMATO)
            except ValueError:
                continue
            return agora() - quando <= timedelta(minutes=VALIDADE_LIBERACAO_MIN)
    return False


def prazo_devolucao(inicio_dt):
    return inicio_dt + timedelta(hours=PRAZO_DEVOLUCAO_HORAS)


# ---------------------------------------------------------------- Transferências
# Formato guardado na coluna K:  "DE > PARA @ 2026-11-09 13:30:00 | DE2 > PARA2 @ ..."
SEP_TRANSF = " | "


def ler_transferencias(texto):
    """Texto da coluna K -> lista de (de, para, datetime ou None)."""
    lista = []
    for parte in str(texto or "").split(SEP_TRANSF):
        if " > " not in parte or " @ " not in parte:
            continue
        nomes, quando = parte.rsplit(" @ ", 1)
        de, para = nomes.split(" > ", 1)
        try:
            dt = datetime.strptime(quando.strip(), FORMATO)
        except ValueError:
            dt = None
        lista.append((de.strip(), para.strip(), dt))
    return lista


def texto_transferencias(texto, motivos=None):
    """Versão legível para a tabela do relatório (com o motivo, se houver)."""
    motivos = motivos or []
    partes = []
    for i, (de, para, dt) in enumerate(ler_transferencias(texto)):
        linha = f"{de} -> {para} ({data_br(dt.strftime(FORMATO)) if dt else '?'})"
        if i < len(motivos) and motivos[i]:
            linha += f" - Motivo: {motivos[i]}"
        partes.append(linha)
    return "; ".join(partes)


def inicio_da_posse(saida_texto, transf_texto):
    """Quando o colaborador ATUAL ficou com o equipamento: na saída, ou na última
    transferência. É daqui que conta o prazo de devolução."""
    _, saida_dt = horas_desde(saida_texto)
    for _, _, dt in reversed(ler_transferencias(transf_texto)):
        if dt is not None:
            return dt
        break
    return saida_dt


def atrasados_do_colaborador(df, colab):
    """Equipamentos que o colaborador está com e não devolveu dentro do prazo (12 horas)."""
    lista = []
    em_uso = df[(df["Status"] == "Em uso") & (df["Colaborador"] == colab)]
    for r in em_uso.itertuples():
        inicio_dt = inicio_da_posse(r.Data_Hora_Saida, r.Transferencias)
        if inicio_dt is not None and agora() > prazo_devolucao(inicio_dt):
            lista.append((r.Equipamento, r.Lacre, prazo_devolucao(inicio_dt)))
    return lista


def html_aviso_pendencia(colab, atrasados):
    itens = "".join(
        '<div class="aviso-item">'
        f"<b>Equipamento:</b> {esc(eq)} - Lacre {esc(lacre)}<br>"
        f"<b>Colaborador:</b> {esc(colab)}<br>"
        f"<b>Prazo de devolução:</b> Vencido (em {prazo:%d/%m/%Y às %H:%M})"
        "</div>"
        for eq, lacre, prazo in atrasados
    )
    return (
        '<div class="aviso"><div class="aviso-tit">⚠️ ATENÇÃO: DEVOLUÇÃO PENDENTE</div>'
        f"<div>O material abaixo ainda não foi devolvido.</div>{itens}</div>"
    )


def registrar_retirada(equip, lacre, colab, assinatura, atual=None):
    """Grava a saída. Devolve o ID novo, ou None se o item já estiver em uso."""
    # Relê a planilha na hora de salvar (evita 2 pessoas pegarem o mesmo item)
    if atual is None:
        atual = carregar_dados()
    ocupado = atual[
        (atual["Status"] == "Em uso")
        & (atual["Equipamento"] == equip) & (atual["Lacre"] == lacre)
    ]
    if not ocupado.empty:
        st.error(f"'{equip}' (lacre {lacre}) já está em uso por {ocupado.iloc[-1]['Colaborador']}.")
        return None
    id_novo = uuid.uuid4().hex[:8]
    aba_registros().append_row(
        [id_novo, equip, str(lacre), colab,
         agora().strftime(FORMATO), assinatura, "", "", "Em uso", 0.0],
        value_input_option="RAW",
    )
    return id_novo


def bloco_autorizacao(p, pend):
    """Aviso + senha do responsável + justificativa para liberar uma nova retirada."""
    ka = st.session_state.aut_key
    st.markdown(html_aviso_pendencia(pend["colab"], pend["atrasados"]), unsafe_allow_html=True)
    st.caption("Nova retirada bloqueada. Para liberar, o responsável informa a senha "
               "e o motivo de o equipamento não ter sido devolvido.")
    senha = st.text_input("Senha do responsável *", type="password", key=f"{p}_senha_{ka}")
    justificativa = st.text_area("Motivo de não ter devolvido *", height=90, key=f"{p}_just_{ka}")

    col_ok, col_cancel = st.columns([1.3, 1])
    with col_ok:
        liberar = st.button("Liberar retirada", type="primary", use_container_width=True,
                            key=f"{p}_lib_{ka}")
    with col_cancel:
        if st.button("Cancelar", use_container_width=True, key=f"{p}_libcancel_{ka}"):
            st.session_state.pop(f"{p}_pend", None)
            st.session_state.aut_key += 1
            st.rerun()

    if not liberar:
        return
    if not senha_correta(senha):
        return
    if not justificativa.strip():
        st.warning("Escreva o motivo de o equipamento não ter sido devolvido.")
        return

    id_novo = registrar_retirada(pend["equip"], pend["lacre"], pend["colab"], pend["assinatura"])
    if not id_novo:
        return
    resumo = "; ".join(f"{eq} (lacre {lc}, venceu em {pz:%d/%m/%Y %H:%M})" for eq, lc, pz in pend["atrasados"])
    registrar_autorizacao(TIPO_RETIRADA, id_novo, pend["colab"], pend["equip"], pend["lacre"],
                          justificativa.strip(), resumo)
    st.session_state.pop(f"{p}_pend", None)
    st.session_state.saida_key += 1
    st.session_state.aut_key += 1
    st.session_state.msg = f"Saída de '{pend['equip']}' registrada para {pend['colab']} com autorização do responsável!"
    st.rerun()


# ---------------------------------------------------------------- Formulários
def form_retirada(p, df, ferramentas, colaboradores, na_janela=False):
    """Formulário de retirada. 'p' é um prefixo para as chaves dos campos."""
    k = st.session_state.saida_key
    em_uso = df[df["Status"] == "Em uso"]
    bloqueados = {chave_equip(r.Equipamento, r.Lacre) for r in em_uso.itertuples()}

    if ferramentas:
        disponiveis = [f for f in ferramentas if chave_equip(f[0], f[1]) not in bloqueados]
        if not disponiveis:
            st.warning("Todos os equipamentos estão em uso.")
        escolha = st.selectbox(
            "Equipamento *", disponiveis, index=None,
            placeholder="Busque pelo nome ou lacre",
            format_func=lambda f: f"{f[0]} | Lacre {f[1]}",
            key=f"{p}_eq_{k}",
        )
        equip, lacre = escolha if escolha else ("", "")
    else:
        equip = st.text_input("Equipamento *", key=f"{p}_eqt_{k}").strip().upper()
        lacre = st.text_input("Lacre", key=f"{p}_lacre_{k}").strip()

    if colaboradores:
        colab = st.selectbox(
            "Colaborador *", colaboradores, index=None,
            placeholder="Selecione o colaborador", key=f"{p}_colab_{k}",
        ) or ""
    else:
        colab = st.text_input("Colaborador *", key=f"{p}_colabt_{k}").strip().upper()

    if na_janela:
        st.text_input("Data e hora da retirada", value=agora().strftime("%d/%m/%Y, %H:%M"),
                      disabled=True, key=f"{p}_dt_{k}")

    st.markdown("**Assinatura do colaborador \\***")
    canvas = campo_assinatura(f"{p}_cv_{k}")
    st.markdown('<div class="ass-legenda">Assinatura digital</div>', unsafe_allow_html=True)

    col_ok, col_cancel = st.columns([1.3, 1]) if na_janela else (st.container(), None)
    with col_ok:
        confirmar = st.button("Confirmar retirada" if na_janela else "Confirmar",
                              type="primary", use_container_width=not na_janela, key=f"{p}_ok_{k}")
    if na_janela and col_cancel is not None:
        with col_cancel:
            if st.button("Cancelar", key=f"{p}_cancel_{k}", use_container_width=True):
                st.rerun()
    if na_janela:
        st.caption("Os campos com * são obrigatórios")

    pend_key = f"{p}_pend"
    if confirmar:
        st.session_state.pop(pend_key, None)
        if not (equip and colab):
            st.warning("Preencha equipamento e colaborador.")
            return
        assinatura = assinatura_para_base64(canvas)
        if assinatura is None:
            aviso_sem_assinatura(canvas)
            return
        atual = carregar_dados()
        atrasados = atrasados_do_colaborador(atual, colab)
        if not atrasados:
            if registrar_retirada(equip, lacre, colab, assinatura, atual=atual):
                st.session_state.saida_key += 1
                st.session_state.msg = f"Saída de '{equip}' registrada para {colab}!"
                st.rerun()
            return
        # Tem devolução vencida (mais de 12 horas): guarda o pedido e pede a senha do responsável
        st.session_state[pend_key] = {
            "equip": equip, "lacre": lacre, "colab": colab,
            "assinatura": assinatura, "atrasados": atrasados,
        }

    pend = st.session_state.get(pend_key)
    if pend and (pend["equip"], pend["lacre"], pend["colab"]) != (equip, lacre, colab):
        st.session_state.pop(pend_key, None)  # mudou equipamento/colaborador: o pedido antigo cai
        pend = None
    if pend:
        bloco_autorizacao(p, pend)


def form_devolucao(p, df, id_fixo=None, na_janela=False):
    """Formulário de devolução. Se 'id_fixo' vier preenchido, devolve aquele registro."""
    k = st.session_state.entrada_key
    em_uso = df[df["Status"] == "Em uso"]
    if em_uso.empty:
        st.info("Nenhum equipamento em uso no momento.")
        return

    if id_fixo:
        linha = em_uso[em_uso["ID"] == id_fixo]
        if linha.empty:
            st.info("Este equipamento já foi devolvido.")
            return
        r = linha.iloc[0]
        st.markdown(
            f"**{esc(r['Equipamento'])}** · Lacre {esc(r['Lacre'])}  \n"
            f"<span class='mut'>{esc(r['Colaborador'])} · saiu em {esc(r['Data_Hora_Saida'])}</span>",
            unsafe_allow_html=True,
        )
        id_reg = id_fixo
    else:
        opcoes = {
            f"{r.Equipamento} | Lacre {r.Lacre} | {r.Colaborador} | saiu {r.Data_Hora_Saida}": r.ID
            for r in em_uso.itertuples()
        }
        rotulo = st.selectbox("Equipamento em uso *", list(opcoes.keys()), index=None,
                              placeholder="Busque pelo nome, lacre ou colaborador",
                              key=f"{p}_eq_{k}")
        id_reg = opcoes.get(rotulo) if rotulo else None

    if not id_reg:
        st.caption("Selecione o equipamento para continuar.")
        return

    # A assinatura só aparece depois que o almoxarifado liberar com a senha
    if not liberacao_valida(id_reg):
        st.markdown(
            '<div class="aviso"><div class="aviso-tit">🔒 AGUARDANDO LIBERAÇÃO DO ALMOXARIFADO</div>'
            "<div>Entregue o equipamento no almoxarifado. Depois que o responsável liberar, "
            "toque em <b>Verificar liberação</b> para assinar.</div></div>",
            unsafe_allow_html=True,
        )
        st.button("Verificar liberação", use_container_width=True, key=f"{p}_ver_{k}")
        return

    st.markdown("**Assinatura do colaborador \\***")
    canvas = campo_assinatura(f"{p}_cv_{k}")
    st.markdown('<div class="ass-legenda">Assinatura digital</div>', unsafe_allow_html=True)

    col_ok, col_cancel = st.columns([1.3, 1]) if na_janela else (st.container(), None)
    with col_ok:
        confirmar = st.button("Confirmar devolução" if na_janela else "Confirmar",
                              type="primary", use_container_width=not na_janela, key=f"{p}_ok_{k}")
    if na_janela and col_cancel is not None:
        with col_cancel:
            if st.button("Cancelar", key=f"{p}_cancel_{k}", use_container_width=True):
                st.rerun()

    if confirmar:
        if not id_reg:
            st.warning("Selecione o equipamento.")
            return
        assinatura = assinatura_para_base64(canvas)
        if assinatura is None:
            aviso_sem_assinatura(canvas)
            return
        if not liberacao_valida(id_reg):
            st.error("A liberação expirou. Peça ao almoxarifado para liberar novamente.")
            return
        ws = aba_registros()
        atual = carregar_dados()
        linha_df = atual[(atual["ID"] == id_reg) & (atual["Status"] == "Em uso")]
        if linha_df.empty:
            st.error("Este equipamento já foi devolvido.")
            return
        saida = datetime.strptime(linha_df.iloc[0]["Data_Hora_Saida"], FORMATO)
        entrada = agora()
        horas = round((entrada - saida).total_seconds() / 3600, 2)
        numero_linha = ws.col_values(1).index(id_reg) + 1
        ws.batch_update(
            [{"range": f"G{numero_linha}:J{numero_linha}",
              "values": [[entrada.strftime(FORMATO), assinatura, "Devolvido", horas]]}],
            value_input_option="RAW",
        )
        st.session_state.entrada_key += 1
        st.session_state.msg = f"Entrada registrada! Tempo de uso: {formata_tempo(horas)}."
        st.rerun()


def garantir_coluna_transferencias(ws):
    """Cria o cabeçalho da coluna K (Transferencias) na aba Registros, se faltar."""
    if ws.col_count < len(COLUNAS):
        ws.add_cols(len(COLUNAS) - ws.col_count)
    cab = ws.row_values(1)
    if len(cab) < len(COLUNAS) or cab[len(COLUNAS) - 1].strip() != COLUNAS[-1]:
        ws.update_cell(1, len(COLUNAS), COLUNAS[-1])


def form_transferencia(p, df, colaboradores, id_fixo=None, na_janela=False):
    """Passa um equipamento em uso para outro colaborador. O novo colaborador
    só assina na devolução; o prazo de 12 h passa a contar a partir da transferência."""
    k = st.session_state.transf_key
    em_uso = df[df["Status"] == "Em uso"]
    if em_uso.empty:
        st.info("Nenhum equipamento em uso no momento.")
        return

    if id_fixo:
        linha = em_uso[em_uso["ID"] == id_fixo]
        if linha.empty:
            st.info("Este equipamento já foi devolvido.")
            return
        r = linha.iloc[0]
        st.markdown(
            f"**{esc(r['Equipamento'])}** · Lacre {esc(r['Lacre'])}  \n"
            f"<span class='mut'>Com {esc(r['Colaborador'])} · saiu em {esc(r['Data_Hora_Saida'])}</span>",
            unsafe_allow_html=True,
        )
        id_reg = id_fixo
    else:
        opcoes = {
            f"{r.Equipamento} | Lacre {r.Lacre} | {r.Colaborador} | saiu {r.Data_Hora_Saida}": r.ID
            for r in em_uso.itertuples()
        }
        rotulo = st.selectbox("Equipamento em uso *", list(opcoes.keys()), index=None,
                              placeholder="Busque pelo nome, lacre ou colaborador",
                              key=f"{p}_eq_{k}")
        id_reg = opcoes.get(rotulo) if rotulo else None
        if not id_reg:
            st.caption("Selecione o equipamento para continuar.")
            return

    atual_colab = em_uso[em_uso["ID"] == id_reg].iloc[0]["Colaborador"]

    if colaboradores:
        novo = st.selectbox(
            "Transferir para *", [c for c in colaboradores if c != atual_colab],
            index=None, placeholder="Selecione o novo colaborador", key=f"{p}_novo_{k}",
        ) or ""
    else:
        novo = st.text_input("Transferir para *", key=f"{p}_novot_{k}").strip().upper()
    motivo = st.text_area("Motivo da transferência (opcional)", height=80, key=f"{p}_mot_{k}")
    st.caption(f"O novo colaborador só assina na devolução. O prazo de {PRAZO_DEVOLUCAO_HORAS} horas "
               "passa a contar a partir de agora.")

    col_ok, col_cancel = st.columns([1.3, 1]) if na_janela else (st.container(), None)
    with col_ok:
        confirmar = st.button("Confirmar transferência", type="primary",
                              use_container_width=True, key=f"{p}_ok_{k}")
    if na_janela and col_cancel is not None:
        with col_cancel:
            if st.button("Cancelar", key=f"{p}_cancel_{k}", use_container_width=True):
                st.rerun()

    if not confirmar:
        return
    if not novo:
        st.warning("Selecione o colaborador que vai receber o equipamento.")
        return
    if novo == atual_colab:
        st.warning("O equipamento já está com este colaborador.")
        return

    ws = aba_registros()
    atual = carregar_dados()  # relê na hora de salvar
    linha_df = atual[(atual["ID"] == id_reg) & (atual["Status"] == "Em uso")]
    if linha_df.empty:
        st.error("Este equipamento já foi devolvido.")
        return
    reg = linha_df.iloc[0]
    de = reg["Colaborador"]
    agora_txt = agora().strftime(FORMATO)
    historico = reg["Transferencias"].strip()
    historico = (historico + SEP_TRANSF if historico else "") + f"{de} > {novo} @ {agora_txt}"

    garantir_coluna_transferencias(ws)
    numero_linha = ws.col_values(1).index(id_reg) + 1
    ws.batch_update(
        [
            {"range": f"D{numero_linha}", "values": [[novo]]},
            {"range": f"K{numero_linha}", "values": [[historico]]},
        ],
        value_input_option="RAW",
    )
    registrar_autorizacao(TIPO_TRANSFERENCIA, id_reg, novo, reg["Equipamento"], reg["Lacre"],
                          motivo.strip(), f"De {de} para {novo}")
    st.session_state.transf_key += 1
    st.session_state.msg = (f"'{reg['Equipamento']}' transferido de {de} para {novo}. "
                            f"{novo} assina na devolução.")
    st.rerun()


@st.dialog("Registrar movimentação")
def janela_movimentacao(df, ferramentas, colaboradores):
    st.caption("Registre a retirada, devolução ou transferência de um equipamento")
    aba_ret, aba_dev, aba_transf = st.tabs(["Retirada", "Devolução", "Transferência"])
    with aba_ret:
        form_retirada("jr", df, ferramentas, colaboradores, na_janela=True)
    with aba_dev:
        form_devolucao("jd", df, na_janela=True)
    with aba_transf:
        form_transferencia("jt", df, colaboradores, na_janela=True)


@st.dialog("Registrar devolução")
def janela_devolucao(id_reg, df):
    form_devolucao("jf", df, id_fixo=id_reg, na_janela=True)


@st.dialog("Transferir equipamento")
def janela_transferencia(id_reg, df, colaboradores):
    form_transferencia("jx", df, colaboradores, id_fixo=id_reg, na_janela=True)


# ---------------------------------------------------------------- Páginas
menu = st.sidebar.selectbox("Navegação", ["Painel", "Relatórios", "Almoxarifado"])

if "msg" in st.session_state:
    st.success(st.session_state.pop("msg"))

if menu == "Painel":
    df = carregar_dados()
    ferramentas = ler_lista("Ferramentas", ["Equipamento", "Lacre"], 2)
    colaboradores = [c[0] for c in ler_lista("Colaboradores", ["Colaborador"], 1)]
    em_uso = df[df["Status"] == "Em uso"]
    bloqueados = {chave_equip(r.Equipamento, r.Lacre) for r in em_uso.itertuples()}

    # ----- topo
    topo_titulo, topo_botao = st.columns([3.5, 1.3], vertical_alignment="center")
    with topo_titulo:
        st.markdown('<p class="tit">Controle de Equipamentos</p>'
                    '<p class="sub">Gerencie retiradas e devoluções em um só lugar</p>',
                    unsafe_allow_html=True)
    with topo_botao:
        if st.button("＋  Registrar movimentação", type="primary", use_container_width=True):
            janela_movimentacao(df, ferramentas, colaboradores)

    st.write("")

    # ----- indicadores
    pendentes = 0
    for r in em_uso.itertuples():
        _, saida_dt = horas_desde(r.Data_Hora_Saida)
        if saida_dt is not None and saida_dt.date() < agora().date():
            pendentes += 1
    if ferramentas:
        disponiveis_qtd = len([f for f in ferramentas if chave_equip(f[0], f[1]) not in bloqueados])
    else:
        disponiveis_qtd = "—"

    k1, k2, k3 = st.columns(3)
    with k1:
        card_indicador(ICONE_OK, "rgba(34,197,94,.2)", "#22c55e", "Disponíveis", disponiveis_qtd)
    with k2:
        card_indicador(ICONE_USO, "rgba(59,130,246,.2)", "#3b82f6", "Em uso", len(em_uso))
    with k3:
        card_indicador(ICONE_ALERTA, "rgba(245,158,11,.2)", "#f59e0b", "Pendentes de devolução", pendentes)

    st.write("")

    # ----- retirada + devoluções
    col_esq, col_dir = st.columns([1, 1.5])

    with col_esq:
        with st.container(key="card_retirada"):
            st.markdown("### Registrar retirada")
            form_retirada("pr", df, ferramentas, colaboradores)

    with col_dir:
        with st.container(key="card_devolucoes"):
            st.markdown("### Devoluções")
            if em_uso.empty:
                st.caption("Nenhum equipamento em uso no momento.")
            else:
                larguras = [1.7, 1.4, 1.1, 1.1, 1.3, 1.3]
                cab = st.columns(larguras)
                for c, nome in zip(cab, ["Equipamento", "Colaborador", "Duração", "Status", "", ""]):
                    c.markdown(f'<div class="th">{nome or "&nbsp;"}</div>', unsafe_allow_html=True)
                for r in em_uso.itertuples():
                    horas, saida_dt = horas_desde(r.Data_Hora_Saida)
                    atrasado = saida_dt is not None and saida_dt.date() < agora().date()
                    c = st.columns(larguras, vertical_alignment="center")
                    c[0].markdown(f"**{esc(r.Equipamento)}**  \n<span class='mut'>Lacre {esc(r.Lacre)}</span>",
                                  unsafe_allow_html=True)
                    transf = ler_transferencias(r.Transferencias)
                    c[1].markdown(
                        esc(r.Colaborador)
                        + (f"  \n<span class='mut'>de {esc(transf[-1][0])}</span>" if transf else ""),
                        unsafe_allow_html=True,
                    )
                    c[2].markdown(formata_tempo(horas))
                    c[3].markdown(
                        '<span class="badge b-warn">Pendente</span>' if atrasado
                        else '<span class="badge b-ok">Em uso</span>',
                        unsafe_allow_html=True,
                    )
                    if c[4].button("Transferir", key=f"transf_{r.Index}_{r.ID}", use_container_width=True):
                        janela_transferencia(r.ID, df, colaboradores)
                    if c[5].button("Devolver", key=f"dev_{r.Index}_{r.ID}", use_container_width=True):
                        janela_devolucao(r.ID, df)
                    st.markdown('<hr class="sep">', unsafe_allow_html=True)

elif menu == "Almoxarifado":
    st.markdown('<p class="tit">Almoxarifado</p>'
                '<p class="sub">Libere a assinatura de devolução para o colaborador</p>',
                unsafe_allow_html=True)
    st.write("")
    df = carregar_dados()
    em_uso = df[df["Status"] == "Em uso"]

    if em_uso.empty:
        st.info("Nenhum equipamento em uso no momento.")
    else:
        kl = st.session_state.lib_key
        with st.container(key="card_liberacao"):
            st.markdown("### Liberar assinatura de devolução")
            opcoes = {
                f"{r.Equipamento} | Lacre {r.Lacre} | {r.Colaborador} | saiu {r.Data_Hora_Saida}": r
                for r in em_uso.itertuples()
            }
            rotulo = st.selectbox("Equipamento em uso *", list(opcoes.keys()), index=None,
                                  placeholder="Busque pelo nome, lacre ou colaborador",
                                  key=f"lib_eq_{kl}")
            senha = st.text_input("Senha do responsável *", type="password", key=f"lib_senha_{kl}")
            st.caption(f"A liberação vale por {VALIDADE_LIBERACAO_MIN} minutos.")
            if st.button("Liberar assinatura", type="primary", key=f"lib_ok_{kl}"):
                if not rotulo:
                    st.warning("Selecione o equipamento.")
                elif senha_correta(senha):
                    r = opcoes[rotulo]
                    registrar_autorizacao(TIPO_LIB_DEVOLUCAO, r.ID, r.Colaborador, r.Equipamento, r.Lacre)
                    st.session_state.lib_key += 1
                    st.session_state.msg = (f"Assinatura liberada para {r.Colaborador} "
                                            f"({r.Equipamento}, lacre {r.Lacre}).")
                    st.rerun()

else:
    st.markdown('<p class="tit">Relatórios</p>'
                '<p class="sub">Consulte as movimentações por dia e baixe o PDF</p>',
                unsafe_allow_html=True)
    st.write("")
    df = carregar_dados()

    if df.empty:
        st.info("Ainda não há registros.")
    else:
        df["_saida"] = pd.to_datetime(df["Data_Hora_Saida"], errors="coerce")
        df["Tempo_Uso_Horas"] = df["Tempo_Uso_Horas"].apply(para_float)

        with st.container(key="card_filtro"):
            periodo = st.date_input(
                "Filtrar por dia (escolha o mesmo dia duas vezes para um único dia)",
                value=(date.today(), date.today()),
                format="DD/MM/YYYY",
            )
        if isinstance(periodo, (tuple, list)):
            inicio = periodo[0]
            fim = periodo[1] if len(periodo) > 1 else periodo[0]
        else:
            inicio = fim = periodo

        filtrado = df[(df["_saida"].dt.date >= inicio) & (df["_saida"].dt.date <= fim)].copy()
        titulo = inicio.strftime("%d/%m/%Y") if inicio == fim else f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y}"

        st.write("")
        motivos = carregar_motivos()
        filtrado["Transferência"] = [
            texto_transferencias(t, motivos["transf"].get(i, []))
            for t, i in zip(filtrado["Transferencias"], filtrado["ID"])
        ]
        filtrado["Retirada em atraso (motivo)"] = [
            motivos["retirada"][i][0] if i in motivos["retirada"] else ""
            for i in filtrado["ID"]
        ]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Movimentações", len(filtrado))
        c2.metric("Em uso", int((filtrado["Status"] == "Em uso").sum()))
        c3.metric("Horas de uso", formata_tempo(filtrado["Tempo_Uso_Horas"].sum()))
        c4.metric("Com transferência", int((filtrado["Transferência"] != "").sum()))
        st.write("")

        if filtrado.empty:
            st.info(f"Nenhuma movimentação em {titulo}.")
        else:
            with st.container(key="card_tabela"):
                tab1, tab2 = st.tabs(["Movimentações", "Resumo por equipamento"])
                with tab1:
                    st.dataframe(
                        filtrado[["Equipamento", "Lacre", "Colaborador", "Data_Hora_Saida",
                                  "Data_Hora_Entrada", "Status", "Tempo_Uso_Horas", "Transferência",
                                  "Retirada em atraso (motivo)"]],
                        use_container_width=True, hide_index=True,
                    )
                with tab2:
                    resumo = (
                        filtrado.groupby(["Equipamento", "Lacre"])["Tempo_Uso_Horas"]
                        .sum().reset_index()
                    )
                    st.dataframe(resumo, use_container_width=True, hide_index=True)

                pdf_bytes = gerar_pdf(filtrado.sort_values("_saida"), titulo, motivos)
                st.download_button(
                    "Baixar relatório em PDF",
                    data=pdf_bytes,
                    file_name=f"relatorio_equipamentos_{inicio:%Y-%m-%d}_{fim:%Y-%m-%d}.pdf",
                    mime="application/pdf",
                    type="primary",
                )
