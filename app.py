import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Controle de Equipamentos", layout="wide")

# Conecta ao Google Sheets usando as credenciais ocultas do Streamlit
@st.cache_resource
def conectar_google_sheets():
    # As credenciais virão das configurações do Streamlit Cloud (Secrets)
    credenciais = dict(st.secrets["gcp_service_account"])
    escopo = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(credenciais, scopes=escopo)
    client = gspread.authorize(creds)
    
    # COLOQUE AQUI O LINK DA SUA PLANILHA DO GOOGLE
    planilha_url = "COLE_AQUI_O_LINK_DA_SUA_PLANILHA" 
    sheet = client.open_by_url(planilha_url).sheet1
    return sheet

sheet = conectar_google_sheets()

def carregar_dados():
    dados = sheet.get_all_records()
    if not dados:
        return pd.DataFrame(columns=['Equipamento', 'Responsavel', 'Data_Hora_Saida', 'Data_Hora_Entrada', 'Status', 'Tempo_Uso_Horas'])
    return pd.DataFrame(dados)

def salvar_dados(df):
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())

df = carregar_dados()

st.title("Sistema de Controle de Equipamentos")
menu = st.sidebar.selectbox("Navegação", ["Registrar Movimentação", "Relatórios"])

if menu == "Registrar Movimentação":
    st.header("Registrar Saída ou Entrada")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Registrar Saída (Retirada)")
        equip_saida = st.text_input("Nome do Equipamento (Saída)")
        resp_saida = st.text_input("Nome do Responsável")
        
        if st.button("Registrar Saída"):
            if equip_saida and resp_saida:
                novo_registro = {
                    'Equipamento': equip_saida,
                    'Responsavel': resp_saida,
                    'Data_Hora_Saida': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Data_Hora_Entrada': '',
                    'Status': 'Em uso',
                    'Tempo_Uso_Horas': 0.0
                }
                df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                salvar_dados(df)
                st.success(f"Saída de '{equip_saida}' registrada!")
            else:
                st.warning("Preencha o equipamento e o responsável.")

    with col2:
        st.subheader("Registrar Entrada (Devolução)")
        equipamentos_em_uso = df[df['Status'] == 'Em uso']['Equipamento'].tolist()
        
        if equipamentos_em_uso:
            equip_entrada = st.selectbox("Selecione o Equipamento", equipamentos_em_uso)
            
            if st.button("Registrar Entrada"):
                agora = datetime.now()
                idx = df[(df['Equipamento'] == equip_entrada) & (df['Status'] == 'Em uso')].index[-1]
                
                df.at[idx, 'Data_Hora_Entrada'] = agora.strftime("%Y-%m-%d %H:%M:%S")
                df.at[idx, 'Status'] = 'Devolvido'
                
                saida = datetime.strptime(df.at[idx, 'Data_Hora_Saida'], "%Y-%m-%d %H:%M:%S")
                tempo_uso = (agora - saida).total_seconds() / 3600
                df.at[idx, 'Tempo_Uso_Horas'] = round(tempo_uso, 2)
                
                salvar_dados(df)
                st.success(f"Entrada de '{equip_entrada}' registrada! Tempo: {round(tempo_uso, 2)}h.")
        else:
            st.info("Nenhum equipamento em uso no momento.")

elif menu == "Relatórios":
    st.header("Painel de Relatórios")
    
    if df.empty or len(df[df['Status'] == 'Devolvido']) == 0:
        st.info("Registre entradas e saídas primeiro.")
    else:
        df_relatorio = df[df['Status'] == 'Devolvido'].copy()
        df_relatorio['Data_Hora_Saida'] = pd.to_datetime(df_relatorio['Data_Hora_Saida'])
        df_relatorio['Data'] = df_relatorio['Data_Hora_Saida'].dt.date
        df_relatorio['Semana'] = df_relatorio['Data_Hora_Saida'].dt.isocalendar().week
        df_relatorio['Mes'] = df_relatorio['Data_Hora_Saida'].dt.strftime('%Y-%m')

        tab1, tab2, tab3 = st.tabs(["Diário", "Semanal", "Mensal"])
        with tab1:
            st.dataframe(df_relatorio.groupby(['Data', 'Equipamento'])['Tempo_Uso_Horas'].sum().reset_index(), use_container_width=True)
        with tab2:
            st.dataframe(df_relatorio.groupby(['Semana', 'Equipamento'])['Tempo_Uso_Horas'].sum().reset_index(), use_container_width=True)
        with tab3:
            st.dataframe(df_relatorio.groupby(['Mes', 'Equipamento'])['Tempo_Uso_Horas'].sum().reset_index(), use_container_width=True)