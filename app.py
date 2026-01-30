import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt

# --- CONFIGURAÇÃO E ESTADO ---
st.set_page_config(page_title="Gerador de Escala Diaconato V5.6", layout="wide")

if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None

# --- FUNÇÕES DE APOIO ---
def obter_primeiro_domingo(ano, mes):
    d = date(ano, mes, 1)
    while d.weekday() != 6: d += timedelta(days=1)
    return d

# --- 1. CARGA DE DADOS ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")
arquivos_historicos = st.sidebar.file_uploader("Suba históricos ou consolidado", type=["csv", "xlsx"], accept_multiple_files=True)

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Processamento de Histórico (Equidade)
    contagem_ceia_historico = {nome: 0 for nome in nomes_membros}
    if arquivos_historicos:
        for arq in arquivos_historicos:
            try:
                df_h = pd.read_csv(arq) if arq.name.endswith('.csv') else pd.read_excel(arq)
                if 'historico_ceia' in df_h.columns:
                    for _, row in df_h.iterrows():
                        if row['Nome'] in contagem_ceia_historico:
                            contagem_ceia_historico[row['Nome']] += row['historico_ceia']
                else:
                    cols = [c for c in df_h.columns if any(x in c for x in ["Santa Ceia", "Ornamentação"])]
                    for col in cols:
                        for cel in df_h[col].dropna().astype(str):
                            for nome in nomes_membros:
                                if nome in cel: contagem_ceia_historico[nome] += 1
            except: continue

    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia_historico)

    # --- 2. CONFIGURAÇÕES ---
    st.sidebar.header("2. Configurações")
    hoje = datetime.now()
    ano = st.sidebar.number_input("Ano", 2025, 2030, hoje.year + (1 if hoje.month == 12 else 0))
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=(hoje.month % 12), format_func=lambda x: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes))

    # --- 3. FÉRIAS / AUSÊNCIAS ---
    st.sidebar.header("3. Férias / Ausências")
    ausencias = st.sidebar.data_editor(
        pd.DataFrame(columns=["Membro", "Início", "Fim"]),
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros),
            "Início": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Fim": st.column_config.DateColumn(format="DD/MM/YYYY")
        }, num_rows="dynamic"
    )

    if st.sidebar.button("Gerar Escala Atualizada"):
        # Lógica de datas (corrigindo formato DD/MM/AAAA)
        data_ini = date(ano, mes, 1)
        data_fim = (date(ano + (1 if mes==12 else 0), 1 if mes==12 else mes+1, 1) - timedelta(days=1))
        datas_mes = pd.date_range(data_ini, data_fim)
        
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        membros_ultimo_culto = []

        for data in datas_mes:
            data_atual = data.date()
            mapa = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
            nome_col = mapa.get(data.weekday())

            if nome_col in dias_semana:
                cands = df_membros[df_membros[nome_col] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # Filtro Ausências
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Início']) and pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                        cands = cands[cands['Nome'] != aus['Membro']]

                # Montagem da linha (Data formatada aqui)
                # %a em PT-BR pode variar, então forçamos o formato limpo
                dias_pt = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
                dia_nome = dias_pt[data.weekday()]
                dia_escala = {"Data": f"{data.strftime('%d/%m/%Y')} ({dia_nome})"}
                
                # Vagas
                vagas = ["Portaria 1", "Portaria 2", "Frente"]
                escalados_dia = []
                for v in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia)].sort_values(by=['historico_ceia', 'escalas_no_mes'])
                    if not v_cands.empty:
                        escolhido = v_cands.iloc[0]['Nome']
                        dia_escala[v] = escolhido
                        escalados_dia.append(escolhido)
                        df_membros.loc[df_membros['Nome'] == escolhido, 'escalas_no_mes'] += 1
                
                # Função Abertura
                aptos_ab = df_membros[(df_membros['Abertura'] == "SIM") & (df_membros['Nome'].isin(cands['Nome']))]
                if not aptos_ab.empty:
                    ab_escolhido = aptos_ab.sort_values(by='escalas_no_mes').iloc[0]['Nome']
                    dia_escala["Abertura"] = ab_escolhido
                
                # Ceia
                if data_atual == data_ceia:
                    dia_escala["Santa Ceia"] = "Equipe Escalada"
                
                escala_final.append(dia_escala)
                membros_ultimo_culto = escalados_dia

        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        st.session_state.df_memoria = df_membros[['Nome', 'historico_ceia']]

    # --- ÁREA DE DOWNLOAD (ONDE O ERRO OCORRIA) ---
    if st.session_state.escala_gerada is not None:
        df_show = st.session_state.escala_gerada
        st.dataframe(df_show, use_container_width=True)

        col_ex, col_img = st.columns(2)
        with col_ex:
            output = io.BytesIO()
            # O SEGREDO: Converter tudo para String antes de escrever no Excel para evitar o TypeError
            df_excel = df_show.astype(str) 
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_excel.to_excel(writer, index=False, sheet_name='Escala')
                workbook = writer.book
                worksheet = writer.sheets['Escala']
                
                # Formatação
                fmt_header = workbook.add_format({'bold': True, 'fg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
                for col_num, value in enumerate(df_excel.columns.values):
                    worksheet.write(0, col_num, value, fmt_header)
                worksheet.set_column(0, 10, 20)

            st.download_button("📥 Baixar Excel", output.getvalue(), "escala.xlsx", key="down_ex")

        with col_img:
            # Geração de Imagem
            fig, ax = plt.subplots(figsize=(10, len(df_show)*0.5 + 1))
            ax.axis('off')
            tab = ax.table(cellText=df_show.values, colLabels=df_show.columns, loc='center', cellLoc='center')
            tab.auto_set_font_size(False)
            tab.set_fontsize(9)
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight')
            st.download_button("📸 Baixar Imagem", buf.getvalue(), "escala.png", key="down_img")

else:
    st.info("Aguardando arquivo membros_master.csv")
