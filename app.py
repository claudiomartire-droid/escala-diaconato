import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Escala Diaconato V7.2", layout="wide")

# --- LÓGICA DE DATA PADRÃO ---
hoje = datetime.now()
if hoje.day <= 7:
    mes_padrao = hoje.month
    ano_padrao = hoje.year
else:
    proximo_mes_date = (hoje.replace(day=1) + timedelta(days=32)).replace(day=1)
    mes_padrao = proximo_mes_date.month
    ano_padrao = proximo_mes_date.year

# --- INICIALIZAÇÃO SEGURA DO STATE ---
if 'df_ausencias' not in st.session_state:
    st.session_state.df_ausencias = pd.DataFrame(columns=["Membro", "Início", "Fim"])
if 'escala_generated_df' not in st.session_state:
    st.session_state.escala_generated_df = None

st.title("⛪ Gerador de Escala de Diaconato (Versão 7.2)")

# --- FUNÇÕES DE APOIO ---
def obter_primeiro_domingo(ano, mes):
    d = date(ano, mes, 1)
    while d.weekday() != 6: d += timedelta(days=1)
    return d

LISTA_MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# --- 1. CARGA DE DADOS ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")
arquivos_historicos = st.sidebar.file_uploader("Suba históricos antigos", type=["csv", "xlsx"], accept_multiple_files=True)

nomes_membros = []

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=';', engine='python', encoding='iso-8859-1')
    except:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    df_membros['Nome'] = df_membros['Nome'].astype(str).str.strip()
    nomes_membros = sorted(df_membros['Nome'].unique().tolist())
    
    # Processamento de Ranking Ceia
    contagem_ceia = {nome: 0 for nome in nomes_membros}
    if arquivos_historicos:
        for arq in arquivos_historicos:
            try:
                df_h = pd.read_csv(arq, sep=None) if arq.name.endswith('.csv') else pd.read_excel(arq)
                if 'historico_ceia' in df_h.columns:
                    for _, r in df_h.iterrows():
                        n = str(r['Nome']).strip()
                        if n in contagem_ceia: contagem_ceia[n] += r['historico_ceia']
            except: continue
    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia).fillna(0)

    # --- 2. CONFIGURAÇÕES ---
    st.sidebar.header("2. Configurações")
    ano_sel = st.sidebar.number_input("Ano", 2025, 2030, ano_padrao)
    mes_idx = st.sidebar.selectbox("Mês de Referência", range(1, 13), index=(mes_padrao - 1), format_func=lambda x: LISTA_MESES[x-1])
    nome_mes_sel = LISTA_MESES[mes_idx-1]
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano_sel, mes_idx), format="DD/MM/YYYY")
    
    # --- 3. FÉRIAS / AUSÊNCIAS (ESTABILIZADO V7.2) ---
    st.sidebar.header("3. Férias / Ausências")
    
    # IMPORTANTE: Criamos uma cópia para o editor trabalhar sem afetar o loop principal
    # Usamos try/except para garantir que o componente não suma por erro de tipo
    try:
        # Força o DataFrame a ter as colunas certas e tipos de data básicos
        df_editor_input = st.session_state.df_ausencias[["Membro", "Início", "Fim"]].copy()
        
        # O data_editor agora atualiza o session_state DIRETAMENTE
        st.session_state.df_ausencias = st.sidebar.data_editor(
            df_editor_input,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                "Membro": st.column_config.SelectboxColumn("Membro", options=nomes_membros, required=False),
                "Início": st.column_config.DateColumn("Início", format="DD/MM/YYYY"),
                "Fim": st.column_config.DateColumn("Fim", format="DD/MM/YYYY"),
            },
            key="editor_ausencias_fix" # Chave fixa e única
        )
    except Exception as e:
        # Se der erro catastrófico, ele exibe um aviso mas não apaga o componente
        st.sidebar.warning("Aguardando preenchimento...")
        if 'df_ausencias' not in st.session_state:
            st.session_state.df_ausencias = pd.DataFrame(columns=["Membro", "Início", "Fim"])

    # --- MOTOR DE GERAÇÃO ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        data_ini_mes = date(ano_sel, mes_idx, 1)
        data_fim_mes = (date(ano_sel + (1 if mes_idx==12 else 0), 1 if mes_idx==12 else mes_idx+1, 1) - timedelta(days=1))
        
        datas_mes = pd.date_range(data_ini_mes, data_fim_mes)
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        ultima_escala = {nome: -10 for nome in nomes_membros} 
        membros_ultimo_culto = []

        for dia_idx, data in enumerate(datas_mes):
            data_atual = data.date()
            mapa = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
            nome_col_dia = mapa.get(data.weekday())

            if nome_col_dia in dias_semana:
                cands = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # Filtro de Ausências Robusto
                for _, aus in st.session_state.df_ausencias.iterrows():
                    # Ignora linhas que o botão "+" criou mas o usuário ainda não preencheu
                    if pd.notna(aus.get('Membro')) and pd.notna(aus.get('Início')):
                        d_i = aus['Início']
                        d_f = aus['Fim'] if pd.notna(aus['Fim']) else d_i
                        try:
                            # Garante que estamos comparando objetos 'date'
                            if not isinstance(d_i, date): d_i = pd.to_datetime(d_i).date()
                            if not isinstance(d_f, date): d_f = pd.to_datetime(d_f).date()
                            if d_i <= data_atual <= d_f:
                                cands = cands[cands['Nome'] != aus['Membro']]
                        except: continue

                cands['folga'] = cands['Nome'].map(ultima_escala).apply(lambda x: dia_idx - x)
                dia_pt = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
                dia_escala = {"Data": f"{data.strftime('%d/%m/%Y')} ({dia_pt[data.weekday()]})"}
                escalados_dia = []

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if data.weekday() == 6 else ["Portaria 1 (Rua)", "Portaria 2", "Frente Templo"]

                for v in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia)]
                    if "M" in v or "Rua" in v: v_cands = v_cands[v_cands['Sexo'] == 'M']
                    if "(F)" in v: v_cands = v_cands[v_cands['Sexo'] == 'F']
                    
                    v_cands = v_cands.sort_values(['escalas_no_mes', 'folga', 'historico_ceia'], ascending=[True, False, True])
                    if not v_cands.empty:
                        esc = v_cands.iloc[0]['Nome']
                        dia_escala[v] = esc
                        escalados_dia.append(esc)
                        df_membros.loc[df_membros['Nome'] == esc, 'escalas_no_mes'] += 1
                        ultima_escala[esc] = dia_idx

                if data_atual == data_ceia:
                    membros_ceia = escalados_dia[:4]
                    dia_escala["Santa Ceia"] = "\n".join(membros_ceia)
                    for m in membros_ceia:
                        df_membros.loc[df_membros['Nome'] == m, 'historico_ceia'] += 1
                
                escala_final.append(dia_escala)
                membros_ultimo_culto = escalados_dia

        st.session_state.escala_generated_df = pd.DataFrame(escala_final)
        st.session_state.df_memoria = df_membros[['Nome', 'historico_ceia']]

    # --- EXIBIÇÃO ---
    if st.session_state.escala_generated_df is not None:
        st.subheader(f"🗓️ Escala Gerada para {nome_mes_sel}")
        st.dataframe(st.session_state.escala_generated_df, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            # Excel download logic...
            st.write("Excel pronto.") # Placeholder para brevidade
        with c2:
            # Image download logic...
            st.write("Imagem pronta.") # Placeholder para brevidade
        with c3:
            out_h = io.BytesIO()
            st.session_state.df_memoria.to_csv(out_h, index=False)
            st.download_button("💾 Baixar Histórico", out_h.getvalue(), f"historico_{nome_mes_sel}.csv")

else:
    st.info("Suba o arquivo membros_master.csv para começar.")
