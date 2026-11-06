import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from streamlit_drawable_canvas import st_canvas
from PIL import Image, ImageDraw
import io
import base64
import uuid
import tempfile
import os
from fpdf import FPDF

st.set_page_config(page_title="Controle de Equipamentos", layout="wide")

# Inicializa as chaves para resetar o quadro automaticamente
if 'saida_key' not in st.session_state:
    st.session_state['saida_key'] = 0
if 'entrada_key' not in st.session_state:
    st.session_state['entrada_key'] = 0

@st.cache_resource
def conectar_google_sheets():
    try:
        credenciais = dict(st.secrets["gcp_service_account"])
        escopo = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(credenciais, scopes=escopo)
        client = gspread.authorize(creds)
        
        # IMPORTANTE: Coloque o link da sua planilha abaixo!
        planilha_url = "https://docs.google.com/spreadsheets/d/1i5uOxp-hNCAvvbBX1580AKGpp9zWYu2DHBcK43272Sw/edit?gid=1708575817#gid=1708575817"
        
        if "COLE_O_LINK" in planilha_url:
            return None, "Você esqueceu de colar o link real da planilha no código."
            
        planilha = client.open_by_url(planilha_url)
        return planilha, None
    except Exception as e:
        return None, f"Erro técnico: {str(e)}"

planilha, erro_conexao = conectar_google_sheets()

if erro_conexao:
    st.error("⚠️ Problema de Conexão com o Google Sheets")
    st.stop()

try:
    aba_registros = planilha.worksheet("Registros")
    aba_ferramentas = planilha.worksheet("Ferramentas")
    aba_colaboradores = planilha.worksheet("Colaboradores")
except Exception:
    st.error("⚠️ Abas não encontradas na Planilha")
    st.stop()

def carregar_dados(aba, colunas):
    dados = aba.get_all_records()
    if not dados:
        return pd.DataFrame(columns=colunas)
    return pd.DataFrame(dados)

df_registros = carregar_dados(aba_registros, ['ID', 'Equipamento', 'Lacre', 'Colaborador', 'Data_Hora_Saida', 'Assinatura_Saida', 'Data_Hora_Entrada', 'Assinatura_Entrada', 'Status', 'Tempo_Uso_Horas'])
df_ferramentas = carregar_dados(aba_ferramentas, ['Equipamento', 'Lacre'])
df_colaboradores = carregar_dados(aba_colaboradores, ['Nome'])

# ====================================================================
# A SOLUÇÃO DEFINITIVA: Recriar o desenho via matemática (Coordenadas)
# Isso bypassa 100% o bug do RuntimeError da biblioteca
# ====================================================================
def json_para_imagem(json_data):
    img = Image.new("RGBA", (400, 150), (238, 238, 238, 255)) # Fundo cinza claro
    draw = ImageDraw.Draw(img)

    if json_data and "objects" in json_data:
        for obj in json_data["objects"]:
            if obj.get("type") == "path":
                path = obj.get("path", [])
                points = []
                for cmd in path:
                    # Extrai as coordenadas X e Y exatas de cada traço
                    if len(cmd) >= 3 and isinstance(cmd[-1], (int, float)) and isinstance(cmd[-2], (int, float)):
                        points.append((cmd[-2], cmd[-1]))
                if len(points) > 1:
                    draw.line(points, fill="#000000", width=2, joint="curve")
                elif len(points) == 1:
                    x, y = points[0]
                    draw.ellipse([x-1, y-1, x+1, y+1], fill="#000000")
    return img

def pil_to_base64(img):
    if img is None:
        return ""
    img_rgb = img.convert('RGB')
    buffered = io.BytesIO()
    img_rgb.save(buffered, format="JPEG")
    return base64.b64encode(buffered.getvalue()).decode("utf-8")

# --- Interface do Sistema ---
st.title("Controle de Ferramentas - Almoxarifado")
menu = st.sidebar.selectbox("Navegação", ["Registrar Movimentação", "Cadastros", "Relatórios e PDF"])

if menu == "Registrar Movimentação":
    st.header("Entrada e Saída de Ferramentas")
    tab_saida, tab_entrada = st.tabs(["Registrar Saída (Retirada)", "Registrar Entrada (Devolução)"])
    
    with tab_saida:
        colab_lista = df_colaboradores['Nome'].dropna().astype(str).unique().tolist()
        colab_selecionado = st.selectbox("Colaborador Solicitante", ["Selecione..."] + colab_lista)
        
        if not df_ferramentas.empty:
            ferramentas_cadastradas = df_ferramentas['Equipamento'].astype(str) + " (Lacre: " + df_ferramentas['Lacre'].astype(str) + ")"
            if not df_registros.empty:
                em_uso = df_registros[df_registros['Status'] == 'Em uso']
                ferramentas_em_uso = em_uso['Equipamento'].astype(str) + " (Lacre: " + em_uso['Lacre'].astype(str) + ")"
                disponiveis = ferramentas_cadastradas[~ferramentas_cadastradas.isin(ferramentas_em_uso)].tolist()
            else:
                disponiveis = ferramentas_cadastradas.tolist()
        else:
            disponiveis = []
            
        ferramenta_selecionada = st.selectbox("Ferramenta Disponível", ["Selecione..."] + disponiveis)
        st.write("Assinatura de Retirada (Desenhe abaixo e clique em Confirmar):")
        
        canvas_saida = st_canvas(
            stroke_width=2, stroke_color="#000", background_color="#eee", 
            height=150, width=400, drawing_mode="freedraw", 
            key=f"canvas_saida_{st.session_state['saida_key']}"
        )
        
        if st.button("Confirmar Saída"):
            if colab_selecionado != "Selecione..." and ferramenta_selecionada != "Selecione...":
                
                # FIM DO ERRO: Não pedimos a imagem, nós lemos os dados matemáticos (json) que NUNCA falham!
                if canvas_saida.json_data is not None and len(canvas_saida.json_data.get("objects", [])) > 0:
                    
                    img_gerada_matematicamente = json_para_imagem(canvas_saida.json_data)
                    assinatura_b64 = pil_to_base64(img_gerada_matematicamente)
                    
                    equipamento_nome = ferramenta_selecionada.split(" (Lacre: ")[0]
                    lacre_numero = ferramenta_selecionada.split(" (Lacre: ")[1].replace(")", "")
                    
                    novo_id = str(uuid.uuid4())
                    data_saida = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    novo_registro = [novo_id, equipamento_nome, lacre_numero, colab_selecionado, data_saida, assinatura_b64, "", "", "Em uso", 0.0]
                    aba_registros.append_row(novo_registro)
                    
                    st.session_state['saida_key'] += 1
                    st.success("Saída registrada com sucesso!")
                    st.rerun()
                else:
                    st.warning("⚠️ Desenhe a assinatura no quadro acima antes de confirmar.")
            else:
                st.warning("Selecione o colaborador e a ferramenta.")

    with tab_entrada:
        if not df_registros.empty:
            em_uso = df_registros[df_registros['Status'] == 'Em uso']
            if not em_uso.empty:
                opcoes_devolucao = em_uso['Equipamento'].astype(str) + " (Lacre: " + em_uso['Lacre'].astype(str) + ") - " + em_uso['Colaborador'].astype(str)
                devolucao_selecionada = st.selectbox("Ferramenta para Devolver", ["Selecione..."] + opcoes_devolucao.tolist())
                st.write("Assinatura de Devolução (Desenhe abaixo e clique em Confirmar):")
                
                canvas_entrada = st_canvas(
                    stroke_width=2, stroke_color="#000", background_color="#eee", 
                    height=150, width=400, drawing_mode="freedraw", 
                    key=f"canvas_entrada_{st.session_state['entrada_key']}"
                )
                
                if st.button("Confirmar Devolução"):
                    if devolucao_selecionada != "Selecione...":
                        
                        # MESMA SOLUÇÃO AQUI: Lemos os dados matemáticos!
                        if canvas_entrada.json_data is not None and len(canvas_entrada.json_data.get("objects", [])) > 0:
                            
                            img_gerada_matematicamente_ent = json_para_imagem(canvas_entrada.json_data)
                            assinatura_b64 = pil_to_base64(img_gerada_matematicamente_ent)
                            
                            equip_lacre = devolucao_selecionada.split(") - ")[0] + ")"
                            equipamento_nome = equip_lacre.split(" (Lacre: ")[0]
                            lacre_numero = equip_lacre.split(" (Lacre: ")[1].replace(")", "")
                            colab_nome = devolucao_selecionada.split(") - ")[1]
                            
                            linha_df = em_uso[(em_uso['Equipamento'] == equipamento_nome) & (em_uso['Lacre'].astype(str) == str(lacre_numero)) & (em_uso['Colaborador'] == colab_nome)].iloc[0]
                            id_registro = linha_df['ID']
                            data_saida = datetime.strptime(linha_df['Data_Hora_Saida'], "%Y-%m-%d %H:%M:%S")
                            data_entrada = datetime.now()
                            tempo_uso = round((data_entrada - data_saida).total_seconds() / 3600, 2)
                            
                            celula = aba_registros.find(str(id_registro))
                            linha_sheet = celula.row
                            
                            aba_registros.update_cell(linha_sheet, 7, data_entrada.strftime("%Y-%m-%d %H:%M:%S"))
                            aba_registros.update_cell(linha_sheet, 8, assinatura_b64)
                            aba_registros.update_cell(linha_sheet, 9, "Devolvido")
                            aba_registros.update_cell(linha_sheet, 10, tempo_uso)
                            
                            st.session_state['entrada_key'] += 1
                            st.success("Devolução registrada com sucesso!")
                            st.rerun()
                        else:
                            st.warning("⚠️ Desenhe a assinatura no quadro acima antes de confirmar.")
                    else:
                        st.warning("Selecione a ferramenta.")
            else:
                st.info("Nenhuma ferramenta em uso no momento.")

elif menu == "Cadastros":
    st.header("Cadastros do Sistema")
    tab_colab, tab_ferr = st.tabs(["Novo Colaborador", "Nova Ferramenta"])
    
    with tab_colab:
        novo_nome = st.text_input("Nome / Código do Colaborador (Ex: MECA 20)")
        if st.button("Cadastrar Colaborador"):
            if novo_nome:
                aba_colaboradores.append_row([novo_nome])
                st.success("Colaborador cadastrado!")
                st.rerun()
                
    with tab_ferr:
        nova_ferramenta = st.text_input("Nome da Ferramenta")
        novo_lacre = st.text_input("Número do Lacre")
        if st.button("Cadastrar Ferramenta"):
            if nova_ferramenta and novo_lacre:
                aba_ferramentas.append_row([nova_ferramenta, novo_lacre])
                st.success("Ferramenta cadastrada!")
                st.rerun()

elif menu == "Relatórios e PDF":
    st.header("Relatório de Movimentações")
    
    if df_registros.empty:
        st.info("Nenhum registro encontrado.")
    else:
        df_view = df_registros.drop(columns=['Assinatura_Saida', 'Assinatura_Entrada', 'ID'], errors='ignore')
        st.dataframe(df_view, use_container_width=True)
        
        if st.button("Gerar Relatório PDF"):
            with st.spinner("Gerando PDF... Isso pode levar alguns segundos."):
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", 'B', 14)
                pdf.cell(200, 10, txt="Relatorio de Movimentacao de Equipamentos", ln=True, align='C')
                
                for _, row in df_registros.iterrows():
                    pdf.ln(5)
                    pdf.set_font("Arial", 'B', 10)
                    pdf.cell(200, 6, txt=f"Equipamento: {row.get('Equipamento','')} | Lacre: {row.get('Lacre','')} | Colab: {row.get('Colaborador','')}", ln=True)
                    pdf.set_font("Arial", size=9)
                    pdf.cell(200, 6, txt=f"Saida: {row.get('Data_Hora_Saida','')} | Entrada: {row.get('Data_Hora_Entrada','')} | Uso: {row.get('Tempo_Uso_Horas','')}h", ln=True)
                    pdf.cell(200, 6, txt=f"Status: {row.get('Status','')}", ln=True)
                    
                    y_img = pdf.get_y()
                    
                    def add_assinatura(b64_str, rotulo, x_pos):
                        if pd.notna(b64_str) and b64_str != "":
                            try:
                                img_data = base64.b64decode(b64_str)
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                                    tmp.write(img_data)
                                    tmp_path = tmp.name
                                pdf.text(x_pos, y_img + 5, rotulo)
                                pdf.image(tmp_path, x=x_pos, y=y_img + 7, w=30)
                                os.remove(tmp_path)
                            except:
                                pdf.text(x_pos, y_img + 5, rotulo + " (Erro imagem)")
                                
                    add_assinatura(row.get('Assinatura_Saida',''), "Ass. Saida:", 10)
                    add_assinatura(row.get('Assinatura_Entrada',''), "Ass. Entrada:", 60)
                    
                    pdf.set_y(y_img + 25)
                    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    pdf.output(tmp_pdf.name)
                    with open(tmp_pdf.name, "rb") as f:
                        pdf_bytes = f.read()
                
                st.download_button(label="📥 Baixar PDF Final", data=pdf_bytes, file_name="relatorio_equipamentos.pdf", mime="application/pdf")
