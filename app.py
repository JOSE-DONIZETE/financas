import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import io
import os

# Configuração da página
st.set_page_config(
    page_title="Controle Financeiro Pessoal",
    page_icon="💰",
    layout="wide"
)

ARQUIVO_DADOS = "dados.csv"

# --- FUNÇÕES PARA SALVAR E CARREGAR DADOS ---
def carregar_dados():
    """Carrega os dados do arquivo CSV se ele existir; caso contrário, cria os dados iniciais."""
    if os.path.exists(ARQUIVO_DADOS):
        try:
            df = pd.read_csv(ARQUIVO_DADOS)
            df["Data"] = pd.to_datetime(df["Data"]).dt.date
            return df
        except Exception as e:
            st.warning("Erro ao carregar o arquivo salvo. Criando nova base de dados.")
    
    # Dados iniciais de padrão caso o arquivo não exista
    dados_iniciais = [
        {"ID": 1, "Data": date(2026, 8, 1), "Tipo": "Receita", "Valor": 5500.00, "Categoria": "Outros", "Conta": "Banco Principal", "Status": "Realizado", "Descrição": "Salário Mensal"},
        {"ID": 2, "Data": date(2026, 8, 2), "Tipo": "Despesa", "Valor": 450.00, "Categoria": "Mercado", "Conta": "Cartão Nubank", "Status": "Realizado", "Descrição": "Compras do Mês"},
        {"ID": 3, "Data": date(2026, 8, 3), "Tipo": "Despesa", "Valor": 120.00, "Categoria": "Transporte", "Conta": "Cartão Itaú", "Status": "Realizado", "Descrição": "Combustível"},
        {"ID": 4, "Data": date(2026, 8, 5), "Tipo": "Despesa", "Valor": 350.00, "Categoria": "Saúde", "Conta": "Banco Principal", "Status": "Realizado", "Descrição": "Consultas / Remédios"},
        {"ID": 5, "Data": date(2026, 8, 12), "Tipo": "Despesa", "Valor": 1200.00, "Categoria": "Casa", "Conta": "Banco Principal", "Status": "Realizado", "Descrição": "Aluguel"},
        {"ID": 6, "Data": date(2026, 8, 25), "Tipo": "Despesa", "Valor": 300.00, "Categoria": "Mercado", "Conta": "Cartão Nubank", "Status": "Pendente", "Descrição": "Supermercado Reposição"},
        {"ID": 7, "Data": date(2026, 8, 28), "Tipo": "Receita", "Valor": 800.00, "Categoria": "Outros", "Conta": "Banco Principal", "Status": "Pendente", "Descrição": "Projeto Freelance"}
    ]
    df = pd.DataFrame(dados_iniciais)
    salvar_dados(df)
    return df

def salvar_dados(df):
    """Salva o DataFrame de lançamentos permanentemente no arquivo CSV."""
    df.to_csv(ARQUIVO_DADOS, index=False)

# --- INICIALIZAÇÃO DO ESTADO DA APLICAÇÃO ---
if "df_lancamentos" not in st.session_state:
    st.session_state.df_lancamentos = carregar_dados()

CATEGORIAS = ["Alimentação", "Saúde", "Casa", "Educação", "Transporte", "Mercado", "Outros"]
CONTAS = ["Banco Principal", "Cartão Nubank", "Cartão Itaú", "Carteira / Dinheiro"]

st.title("💰 Controle Financeiro Pessoal")

# --- BARRA LATERAL: GERENCIAR E IMPORTAR/EXPORTAR ---
st.sidebar.header("⚙️ Opções & Backup")

# Importação de Arquivo
uploaded_file = st.sidebar.file_uploader("Importar Excel ou CSV", type=["xlsx", "csv"])
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_imp = pd.read_csv(uploaded_file)
        else:
            df_imp = pd.read_excel(uploaded_file)
        
        df_imp["Data"] = pd.to_datetime(df_imp["Data"]).dt.date
        st.session_state.df_lancamentos = df_imp
        salvar_dados(df_imp)
        st.sidebar.success("Dados importados e salvos com sucesso!")
    except Exception as e:
        st.sidebar.error(f"Erro ao carregar arquivo: {e}")

# Exportação de Arquivo
st.sidebar.subheader("📥 Exportar Dados")
col_exp1, col_exp2 = st.sidebar.columns(2)

csv_data = st.session_state.df_lancamentos.to_csv(index=False).encode('utf-8')
col_exp1.download_button("Exportar CSV", data=csv_data, file_name="financeiro.csv", mime="text/csv")

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    st.session_state.df_lancamentos.to_excel(writer, index=False, sheet_name="Lancamentos")
excel_data = buffer.getvalue()
col_exp2.download_button("Exportar Excel", data=excel_data, file_name="financeiro.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.sidebar.markdown("---")

# --- FILTROS DE DATA (Formato BR) ---
st.sidebar.header("🔍 Filtro por Período")
data_inicio = st.sidebar.date_input("Data Inicial", value=date(2026, 8, 1), format="DD/MM/YYYY")
data_fim = st.sidebar.date_input("Data Final", value=date(2026, 8, 31), format="DD/MM/YYYY")

df = st.session_state.df_lancamentos.copy()
if not df.empty:
    df["Data"] = pd.to_datetime(df["Data"]).dt.date
    mask = (df["Data"] >= data_inicio) & (df["Data"] <= data_fim)
    df_filtrado = df[mask].copy()
else:
    df_filtrado = df.copy()

# --- ABA DE NAVEGAÇÃO PRINCIPAL ---
tab_dash, tab_lancamentos, tab_form = st.tabs(["📊 Dashboard & Projeções", "📋 Lista de Lançamentos", "➕ Novo / Editar Lançamento"])

# -----------------------------------------------------------------------------
# TAB 1: DASHBOARD
# -----------------------------------------------------------------------------
with tab_dash:
    st.subheader("Resumo Financeiro do Período Selecionado")
    
    receitas_realizadas = df_filtrado[(df_filtrado["Tipo"] == "Receita") & (df_filtrado["Status"] == "Realizado")]["Valor"].sum()
    despesas_realizadas = df_filtrado[(df_filtrado["Tipo"] == "Despesa") & (df_filtrado["Status"] == "Realizado")]["Valor"].sum()
    saldo_atual = receitas_realizadas - despesas_realizadas

    receitas_pendentes = df_filtrado[(df_filtrado["Tipo"] == "Receita") & (df_filtrado["Status"] == "Pendente")]["Valor"].sum()
    despesas_pendentes = df_filtrado[(df_filtrado["Tipo"] == "Despesa") & (df_filtrado["Status"] == "Pendente")]["Valor"].sum()
    projecao_futura = saldo_atual + (receitas_pendentes - despesas_pendentes)

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Receitas (Realizadas)", f"R$ {receitas_realizadas:,.2f}")
    kpi2.metric("Despesas (Realizadas)", f"R$ {despesas_realizadas:,.2f}")
    kpi3.metric("Saldo Realizado", f"R$ {saldo_atual:,.2f}", delta=f"{saldo_atual:,.2f}")
    kpi4.metric("Projeção Futura", f"R$ {projecao_futura:,.2f}", help="Considera saldo atual + lançamentos pendentes no período.")

    st.markdown("---")
    
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.write("### Despesas por Categoria")
        df_despesas = df_filtrado[df_filtrado["Tipo"] == "Despesa"]
        if not df_despesas.empty:
            fig_pie = px.pie(
                df_despesas, 
                names="Categoria", 
                values="Valor", 
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Nenhuma despesa registrada no período.")

    with col_g2:
        st.write("### Saldo Realizado por Conta/Banco")
        if not df_filtrado.empty:
            df_contas_in = df_filtrado[(df_filtrado["Tipo"] == "Receita") & (df_filtrado["Status"] == "Realizado")].groupby("Conta")["Valor"].sum()
            df_contas_out = df_filtrado[(df_filtrado["Tipo"] == "Despesa") & (df_filtrado["Status"] == "Realizado")].groupby("Conta")["Valor"].sum()
            
            df_saldos = pd.DataFrame({"Entradas": df_contas_in, "Saídas": df_contas_out}).fillna(0)
            df_saldos["Saldo"] = df_saldos["Entradas"] - df_saldos["Saídas"]
            df_saldos = df_saldos.reset_index()

            fig_bar = px.bar(
                df_saldos, 
                x="Conta", 
                y="Saldo", 
                text_auto='.2f', 
                color="Saldo", 
                color_continuous_scale="Blues"
            )
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Sem dados para exibições por conta.")

# -----------------------------------------------------------------------------
# TAB 2: GERENCIAMENTO DE LANÇAMENTOS
# -----------------------------------------------------------------------------
with tab_lancamentos:
    st.subheader("Todos os Lançamentos no Período")
    
    if not df_filtrado.empty:
        df_display = df_filtrado.copy()
        df_display["Data"] = pd.to_datetime(df_display["Data"]).dt.strftime('%d/%m/%Y')
        
        st.dataframe(
            df_display.style.format({"Valor": "R$ {:,.2f}"}),
            use_container_width=True,
            height=350
        )
        
        st.markdown("---")
        st.write("### 🗑️ Excluir Lançamento")
        col_del1, col_del2 = st.columns([3, 1])
        id_excluir = col_del1.selectbox("Selecione o ID do Lançamento para Excluir:", df_filtrado["ID"].unique())
        
        if col_del2.button("Excluir Lançamento", type="primary"):
            st.session_state.df_lancamentos = st.session_state.df_lancamentos[
                st.session_state.df_lancamentos["ID"] != id_excluir
            ]
            salvar_dados(st.session_state.df_lancamentos)  # Salva no disco ao excluir
            st.success(f"Lançamento ID {id_excluir} excluído com sucesso!")
            st.rerun()
    else:
        st.info("Nenhum lançamento encontrado para os filtros selecionados.")

# -----------------------------------------------------------------------------
# TAB 3: NOVO LANÇAMENTO OU EDIÇÃO
# -----------------------------------------------------------------------------
with tab_form:
    st.subheader("Adicionar ou Editar Lançamento")
    
    modo = st.radio("Selecione a ação:", ["Novo Lançamento", "Editar Existente"], horizontal=True, key="modo_acao")
    
    id_edit = None
    dados_edit = {
        "Data": date.today(), 
        "Tipo": "Despesa", 
        "Valor": 0.0, 
        "Categoria": "Alimentação", 
        "Conta": "Banco Principal", 
        "Status": "Realizado", 
        "Descrição": ""
    }
    
    if modo == "Editar Existente" and not st.session_state.df_lancamentos.empty:
        id_edit = st.selectbox("Escolha o ID para editar:", st.session_state.df_lancamentos["ID"].unique(), key="select_id_edit")
        reg = st.session_state.df_lancamentos[st.session_state.df_lancamentos["ID"] == id_edit].iloc[0]
        dados_edit = {
            "Data": reg["Data"],
            "Tipo": reg["Tipo"],
            "Valor": float(reg["Valor"]),
            "Categoria": reg["Categoria"],
            "Conta": reg["Conta"],
            "Status": reg["Status"],
            "Descrição": reg["Descrição"]
        }

    chave_prefix = f"edit_{id_edit}" if modo == "Editar Existente" else "novo_lanc"

    with st.form("form_lancamento"):
        col_f1, col_f2, col_f3 = st.columns(3)
        data_f = col_f1.date_input("Data", value=dados_edit["Data"], format="DD/MM/YYYY", key=f"{chave_prefix}_data")
        tipo_f = col_f2.selectbox("Tipo", ["Receita", "Despesa"], index=0 if dados_edit["Tipo"] == "Receita" else 1, key=f"{chave_prefix}_tipo")
        valor_f = col_f3.number_input("Valor (R$)", value=dados_edit["Valor"], step=10.0, format="%.2f", key=f"{chave_prefix}_valor")

        col_f4, col_f5, col_f6 = st.columns(3)
        cat_index = CATEGORIAS.index(dados_edit["Categoria"]) if dados_edit["Categoria"] in CATEGORIAS else 0
        cat_f = col_f4.selectbox("Categoria", CATEGORIAS, index=cat_index, key=f"{chave_prefix}_cat")
        
        cta_index = CONTAS.index(dados_edit["Conta"]) if dados_edit["Conta"] in CONTAS else 0
        conta_f = col_f5.selectbox("Conta / Cartão", CONTAS, index=cta_index, key=f"{chave_prefix}_cta")
        
        status_f = col_f6.selectbox("Status", ["Realizado", "Pendente"], index=0 if dados_edit["Status"] == "Realizado" else 1, key=f"{chave_prefix}_status")

        desc_f = st.text_input("Descrição / Detalhes", value=dados_edit["Descrição"], key=f"{chave_prefix}_desc")

        btn_salvar = st.form_submit_button("Salvar Lançamento")

        if btn_salvar:
            if valor_f <= 0:
                st.error("⚠️ O valor precisa ser maior que zero (R$ 0,00).")
            else:
                if modo == "Novo Lançamento":
                    novo_id = int(st.session_state.df_lancamentos["ID"].max() + 1) if not st.session_state.df_lancamentos.empty else 1
                    novo_registro = pd.DataFrame([{
                        "ID": novo_id,
                        "Data": data_f,
                        "Tipo": tipo_f,
                        "Valor": valor_f,
                        "Categoria": cat_f,
                        "Conta": conta_f,
                        "Status": status_f,
                        "Descrição": desc_f
                    }])
                    st.session_state.df_lancamentos = pd.concat([st.session_state.df_lancamentos, novo_registro], ignore_index=True)
                    salvar_dados(st.session_state.df_lancamentos)  # Salva no disco
                    st.success("Lançamento adicionado com sucesso!")
                else:
                    idx = st.session_state.df_lancamentos[st.session_state.df_lancamentos["ID"] == id_edit].index[0]
                    st.session_state.df_lancamentos.loc[idx] = [id_edit, data_f, tipo_f, valor_f, cat_f, conta_f, status_f, desc_f]
                    salvar_dados(st.session_state.df_lancamentos)  # Salva no disco
                    st.success(f"Lançamento ID {id_edit} atualizado com sucesso!")
                st.rerun()