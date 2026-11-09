import base64
import io
import os
import tempfile
import uuid
from datetime import datetime, date
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

# Ordem EXATA das colunas da aba "Registros" (A até J)
COLUNAS = [
    "ID", "Equipamento", "Lacre", "Colaborador",
    "Data_Hora_Saida", "Assinatura_Saida",
    "Data_Hora_Entrada", "Assinatura_Entrada",
    "Status", "Tempo_Uso_Horas",
]

# Cole o link da planilha aqui OU coloque em Secrets: planilha_url = "https://..."
PLANILHA_URL = st.secrets.get("planilha_url", "https://docs.google.com/spreadsheets/d/1i5uOxp-hNCAvvbBX1580AKGpp9zWYu2DHBcK43272Sw/edit?gid=0#gid=0")

# Limite de uma célula do Google Sheets = 50.000 caracteres
LIMITE_CELULA = 45000

# Chaves para limpar o quadro de assinatura depois de salvar
for chave in ("saida_key", "entrada_key"):
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


def _num(d, chave, padrao):
    try:
        v = d.get(chave, padrao)
        return padrao if v is None else float(v)
    except (TypeError, ValueError):
        return padrao


def _pontos_do_traco(obj):
    """Converte um 'path' do canvas (fabric.js) em pontos na posição real do quadro."""
    comandos = obj.get("path") or []
    esc_x, esc_y = _num(obj, "scaleX", 1.0), _num(obj, "scaleY", 1.0)
    off = obj.get("pathOffset") or {}
    off_x, off_y = _num(off, "x", 0.0), _num(off, "y", 0.0)
    # centro do objeto no quadro
    cx = _num(obj, "left", 0.0) + _num(obj, "width", 0.0) * esc_x / 2
    cy = _num(obj, "top", 0.0) + _num(obj, "height", 0.0) * esc_y / 2

    def pos(x, y):
        return (cx + (x - off_x) * esc_x, cy + (y - off_y) * esc_y)

    pontos, atual = [], None
    for c in comandos:
        if not c:
            continue
        tipo, v = c[0], [float(n) for n in c[1:]]
        if tipo == "M" and len(v) >= 2:
            atual = pos(v[0], v[1])
            pontos.append(atual)
        elif tipo == "L" and len(v) >= 2:
            atual = pos(v[0], v[1])
            pontos.append(atual)
        elif tipo == "Q" and len(v) >= 4 and atual is not None:
            c1, fim = pos(v[0], v[1]), pos(v[2], v[3])
            for i in range(1, 9):  # curva suave com 8 segmentos
                t = i / 8
                x = (1 - t) ** 2 * atual[0] + 2 * (1 - t) * t * c1[0] + t ** 2 * fim[0]
                y = (1 - t) ** 2 * atual[1] + 2 * (1 - t) * t * c1[1] + t ** 2 * fim[1]
                pontos.append((x, y))
            atual = fim
    return pontos


def assinatura_para_base64(resultado):
    """Redesenha a assinatura a partir do traço do canvas (json_data) e devolve
    PNG pequeno em base64. Retorna None se o quadro estiver vazio."""
    if resultado is None:
        return None
    dados = getattr(resultado, "json_data", None)
    objetos = (dados or {}).get("objects") or []
    tracos = [o for o in objetos if o.get("type") == "path"]
    if not tracos:
        return None

    img = Image.new("L", (LARGURA_ASS * ESCALA, ALTURA_ASS * ESCALA), 255)
    draw = ImageDraw.Draw(img)
    desenhou = False
    for obj in tracos:
        pts = [(x * ESCALA, y * ESCALA) for x, y in _pontos_do_traco(obj)]
        if not pts:
            continue
        largura = max(2, int(_num(obj, "strokeWidth", 2.0) * ESCALA))
        if len(pts) == 1:
            x, y = pts[0]
            draw.ellipse([x - largura / 2, y - largura / 2, x + largura / 2, y + largura / 2], fill=0)
        else:
            draw.line(pts, fill=0, width=largura, joint="curve")
        desenhou = True
    if not desenhou:
        return None

    img = img.resize((LARGURA_ASS, ALTURA_ASS), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    texto = base64.b64encode(buf.getvalue()).decode("ascii")
    return texto if len(texto) <= LIMITE_CELULA else None


def campo_assinatura(chave):
    return st_canvas(
        fill_color="rgba(255,255,255,0)",
        stroke_width=2,
        stroke_color="#000000",
        background_color="#ffffff",
        height=120,
        width=300,
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


def desenha_assinatura(pdf, b64, x, y):
    pdf.rect(x, y, 50, 20)
    if not b64:
        return
    try:
        dados = base64.b64decode(b64)
        img = Image.open(io.BytesIO(dados)).convert("RGB")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            caminho = tmp.name
        img.save(caminho, format="PNG")
        pdf.image(caminho, x=x, y=y, w=50, h=20)
        os.remove(caminho)
    except Exception:
        pdf.set_xy(x + 2, y + 8)
        pdf.set_font("Helvetica", "I", 8)
        pdf.cell(46, 4, "assinatura invalida")


def gerar_pdf(df, titulo_periodo):
    pdf = FPDF()
    pdf.set_auto_page_break(False)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 15)
    pdf.cell(0, 10, limpa("Relatorio de Movimentacao de Equipamentos"), ln=1, align="C")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, limpa(f"Periodo: {titulo_periodo}"), ln=1, align="C")
    pdf.ln(4)

    for _, r in df.iterrows():
        if pdf.get_y() > 235:
            pdf.add_page()
        pdf.set_x(10)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, limpa(f"Equipamento: {r['Equipamento']} | Lacre: {r['Lacre']} | Colab: {r['Colaborador']}"), ln=1)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, limpa(
            f"Saida: {r['Data_Hora_Saida']} | Entrada: {r['Data_Hora_Entrada']} | Uso: {formata_tempo(r['Tempo_Uso_Horas'])}"
        ), ln=1)
        pdf.cell(0, 6, limpa(f"Status: {r['Status']}"), ln=1)
        pdf.ln(1)

        y = pdf.get_y()
        pdf.set_xy(10, y)
        pdf.cell(60, 5, "Ass. Saida:")
        pdf.set_xy(80, y)
        pdf.cell(60, 5, "Ass. Entrada:")
        desenha_assinatura(pdf, r["Assinatura_Saida"], 10, y + 6)
        if r["Status"] == "Devolvido":
            desenha_assinatura(pdf, r["Assinatura_Entrada"], 80, y + 6)
        pdf.set_y(y + 29)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(4)

    saida = pdf.output(dest="S")
    if isinstance(saida, str):
        saida = saida.encode("latin-1")
    return bytes(saida)


# ---------------------------------------------------------------- Interface
st.title("Sistema de Controle de Equipamentos")
menu = st.sidebar.selectbox("Navegação", ["Registrar Movimentação", "Relatórios"])

if "msg" in st.session_state:
    st.success(st.session_state.pop("msg"))

if menu == "Registrar Movimentação":
    df = carregar_dados()
    em_uso = df[df["Status"] == "Em uso"]
    bloqueados = {chave_equip(r.Equipamento, r.Lacre) for r in em_uso.itertuples()}

    col1, col2 = st.columns(2)

    # ---------------- SAÍDA
    with col1:
        st.subheader("Registrar Saída (Retirada)")

        ferramentas = ler_lista("Ferramentas", ["Equipamento", "Lacre"], 2)
        disponiveis = [f for f in ferramentas if chave_equip(f[0], f[1]) not in bloqueados]

        if ferramentas:
            if disponiveis:
                escolha = st.selectbox(
                    "Equipamento disponível",
                    disponiveis,
                    format_func=lambda f: f"{f[0]} | Lacre {f[1]}",
                    key=f"equip_saida_{st.session_state.saida_key}",
                )
                equip_saida, lacre_saida = escolha
            else:
                st.warning("Todos os equipamentos estão em uso.")
                equip_saida = lacre_saida = ""
        else:
            equip_saida = st.text_input("Equipamento", key=f"eq_txt_{st.session_state.saida_key}").strip().upper()
            lacre_saida = st.text_input("Lacre", key=f"lacre_txt_{st.session_state.saida_key}").strip()

        colaboradores = [c[0] for c in ler_lista("Colaboradores", ["Colaborador"], 1)]
        if colaboradores:
            colab_saida = st.selectbox("Colaborador", colaboradores, key=f"colab_saida_{st.session_state.saida_key}")
        else:
            colab_saida = st.text_input("Colaborador", key=f"colab_txt_{st.session_state.saida_key}").strip().upper()

        st.write("Assinatura de saída:")
        canvas_saida = campo_assinatura(f"canvas_saida_{st.session_state.saida_key}")

        if st.button("Registrar Saída", type="primary"):
            if not (equip_saida and colab_saida):
                st.warning("Preencha equipamento e colaborador.")
            else:
                assinatura = assinatura_para_base64(canvas_saida)
                if assinatura is None:
                    st.warning("Assine no quadro antes de registrar.")
                else:
                    # Relê a planilha na hora de salvar (evita 2 pessoas pegarem o mesmo item)
                    atual = carregar_dados()
                    ocupado = atual[
                        (atual["Status"] == "Em uso")
                        & (atual["Equipamento"] == equip_saida)
                        & (atual["Lacre"] == lacre_saida)
                    ]
                    if not ocupado.empty:
                        st.error(f"'{equip_saida}' (lacre {lacre_saida}) já está em uso por {ocupado.iloc[-1]['Colaborador']}.")
                    else:
                        aba_registros().append_row(
                            [
                                uuid.uuid4().hex[:8], equip_saida, str(lacre_saida), colab_saida,
                                agora().strftime(FORMATO), assinatura,
                                "", "", "Em uso", 0.0,
                            ],
                            value_input_option="RAW",
                        )
                        st.session_state.saida_key += 1
                        st.session_state.msg = f"Saída de '{equip_saida}' registrada para {colab_saida}!"
                        st.rerun()

    # ---------------- ENTRADA
    with col2:
        st.subheader("Registrar Entrada (Devolução)")
        if em_uso.empty:
            st.info("Nenhum equipamento em uso no momento.")
        else:
            opcoes = {
                f"{r.Equipamento} | Lacre {r.Lacre} | {r.Colaborador} | saiu {r.Data_Hora_Saida}": r.ID
                for r in em_uso.itertuples()
            }
            rotulo = st.selectbox("Equipamento em uso", list(opcoes.keys()), key=f"equip_entrada_{st.session_state.entrada_key}")

            st.write("Assinatura de entrada:")
            canvas_entrada = campo_assinatura(f"canvas_entrada_{st.session_state.entrada_key}")

            if st.button("Registrar Entrada", type="primary"):
                assinatura = assinatura_para_base64(canvas_entrada)
                if assinatura is None:
                    st.warning("Assine no quadro antes de registrar.")
                else:
                    id_reg = opcoes[rotulo]
                    ws = aba_registros()
                    atual = carregar_dados()
                    linha_df = atual[(atual["ID"] == id_reg) & (atual["Status"] == "Em uso")]
                    if linha_df.empty:
                        st.error("Este equipamento já foi devolvido.")
                    else:
                        saida = datetime.strptime(linha_df.iloc[0]["Data_Hora_Saida"], FORMATO)
                        entrada = agora()
                        horas = round((entrada - saida).total_seconds() / 3600, 2)
                        numero_linha = ws.col_values(1).index(id_reg) + 1
                        ws.batch_update(
                            [{
                                "range": f"G{numero_linha}:J{numero_linha}",
                                "values": [[entrada.strftime(FORMATO), assinatura, "Devolvido", horas]],
                            }],
                            value_input_option="RAW",
                        )
                        st.session_state.entrada_key += 1
                        st.session_state.msg = f"Entrada registrada! Tempo de uso: {formata_tempo(horas)}."
                        st.rerun()

    st.divider()
    st.subheader("Equipamentos em uso agora")
    if em_uso.empty:
        st.caption("Nenhum.")
    else:
        st.dataframe(
            em_uso[["Equipamento", "Lacre", "Colaborador", "Data_Hora_Saida"]],
            use_container_width=True, hide_index=True,
        )

else:
    st.header("Painel de Relatórios")
    df = carregar_dados()

    if df.empty:
        st.info("Ainda não há registros.")
    else:
        df["_saida"] = pd.to_datetime(df["Data_Hora_Saida"], errors="coerce")
        df["Tempo_Uso_Horas"] = df["Tempo_Uso_Horas"].apply(para_float)

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

        c1, c2, c3 = st.columns(3)
        c1.metric("Movimentações", len(filtrado))
        c2.metric("Em uso", int((filtrado["Status"] == "Em uso").sum()))
        c3.metric("Horas de uso", formata_tempo(filtrado["Tempo_Uso_Horas"].sum()))

        if filtrado.empty:
            st.info(f"Nenhuma movimentação em {titulo}.")
        else:
            tab1, tab2 = st.tabs(["Movimentações", "Resumo por equipamento"])
            with tab1:
                st.dataframe(
                    filtrado[["Equipamento", "Lacre", "Colaborador", "Data_Hora_Saida",
                              "Data_Hora_Entrada", "Status", "Tempo_Uso_Horas"]],
                    use_container_width=True, hide_index=True,
                )
            with tab2:
                resumo = (
                    filtrado.groupby(["Equipamento", "Lacre"])["Tempo_Uso_Horas"]
                    .sum().reset_index()
                )
                st.dataframe(resumo, use_container_width=True, hide_index=True)

            pdf_bytes = gerar_pdf(filtrado.sort_values("_saida"), titulo)
            st.download_button(
                "Baixar relatório em PDF",
                data=pdf_bytes,
                file_name=f"relatorio_equipamentos_{inicio:%Y-%m-%d}_{fim:%Y-%m-%d}.pdf",
                mime="application/pdf",
            )
