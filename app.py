import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Gerador de Escala Diaconato V5.7", layout="wide")

if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.7)")

# --- 1. CARGA DE DADOS ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")
arquivos_historicos = st.sidebar.file_uploader("Suba históricos", type=["csv", "xlsx"], accept_multiple_files=True)

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Processamento de Equidade
    contagem_ceia = {nome: 0 for nome in nomes_membros}
    if arquivos_historicos:
        for arq in arquivos_historicos:
            try:
                df_h = pd.read_csv(arq) if arq.name.endswith('.csv') else pd.read_excel(arq)
                if 'historico_ceia' in df_h.columns:
                    for _, r in df_h.iterrows():
                        if r['Nome'] in contagem_ceia: contagem_ceia[r['Nome']] += r['historico_ceia']
            except: continue
    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia)

    # Tabs de Conferência
    st.subheader("📋 Conferência de Regras e Equidade")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Ranking Santa Ceia"])
    with t3: st.dataframe(df_membros[['Nome', 'historico_ceia']].sort_values('historico_ceia'), use_container_width=True)

    # --- 2. CONFIGURAÇÕES ---
    st.sidebar.header("2. Configurações")
    hoje = datetime.now()
    ano = st.sidebar.number_input("Ano", 2025, 2030, hoje.year + (1 if hoje.month == 12 else 0))
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=(hoje.month % 12), format_func=lambda x: ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"][x-1])
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=date(ano, mes, 1))
    
    data_ini = date(ano, mes, 1)
    data_fim = (date(ano + (1 if mes==12 else 0), 1 if mes==12 else mes+1, 1) - timedelta(days=1))
    datas_excluir = st.sidebar.multiselect("Datas EXCLUÍDAS", options=pd.date_range(data_ini, data_fim), format_func=lambda x: x.strftime('%d/%m/%Y'))

    # --- 3. FÉRIAS ---
    st.sidebar.header("3. Férias / Ausências")
    ausencias = st.sidebar.data_editor(pd.DataFrame(columns=["Membro", "Início", "Fim"]), column_config={"Membro": st.column_config.SelectboxColumn(options=nomes_membros), "Início": st.column_config.DateColumn(format="DD/MM/YYYY"), "Fim": st.column_config.DateColumn(format="DD/MM/YYYY")}, num_rows="dynamic")

    # --- MOTOR ---
    if st.sidebar.button("Gerar Escala Completa V5.7"):
        datas_mes = pd.date_range(data_ini, data_fim)
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        membros_ultimo_culto = []

        for data in datas_mes:
            data_atual = data.date()
            if any(data_atual == d.date() for d in datas_excluir): continue
            
            mapa = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
            nome_col = mapa.get(data.weekday())

            if nome_col in dias_semana:
                cands = df_membros[df_membros[nome_col] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # Filtro Ausências
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Início']) and pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                        cands = cands[cands['Nome'] != aus['Membro']]

                dias_pt = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
                dia_escala = {"Data": f"{data.strftime('%d/%m/%Y')} ({dias_pt[data.weekday()]})"}
                escalados_dia = []

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente (M)", "Frente (F)"] if data.weekday() == 6 else ["Portaria 1 (Rua)", "Portaria 2", "Frente"]

                for v in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia)]
                    if "M" in v or "Rua" in v: v_cands = v_cands[v_cands['Sexo'] == 'M']
                    if "F" in v: v_cands = v_cands[v_cands['Sexo'] == 'F']
                    
                    v_cands = v_cands.sort_values(['historico_ceia', 'escalas_no_mes'])
                    if not v_cands.empty:
                        escolhido = v_cands.iloc[0]['Nome']
                        dia_escala[v] = escolhido
                        escalados_dia.append(escolhido)
                        df_membros.loc[df_membros['Nome'] == escolhido, 'escalas_no_mes'] += 1

                # ABERTURA (RESTAURADO)
                aptos_ab = cands[cands['Abertura'] == "SIM"].copy()
                ja_no_dia = [n for n in escalados_dia if n in aptos_ab['Nome'].values]
                if ja_no_dia:
                    dia_escala["Abertura"] = ja_no_dia[0]
                else:
                    sobra_ab = aptos_ab[~aptos_ab['Nome'].isin(escalados_dia)]
                    if not sobra_ab.empty:
                        escolhido_ab = sobra_ab.sort_values(['historico_ceia', 'escalas_no_mes']).iloc[0]['Nome']
                        dia_escala["Abertura"] = escolhido_ab
                        df_membros.loc[df_membros['Nome'] == escolhido_ab, 'escalas_no_mes'] += 0.5
                        escalados_dia.append(escolhido_ab)

                if data_atual == data_ceia:
                    dia_escala["Servir Ceia"] = ", ".join(escalados_dia[:4])
                
                escala_final.append(dia_escala)
                membros_ultimo_culto = escalados_dia

        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        st.session_state.df_memoria = df_membros[['Nome', 'historico_ceia']]

    # --- EXPORTAÇÃO FORMATADA ---
    if st.session_state.escala_gerada is not None:
        st.dataframe(st.session_state.escala_gerada, use_container_width=True)
        
        c_ex, c_img = st.columns(2)
        with c_ex:
            output = io.BytesIO()
            df_ex = st.session_state.escala_gerada.fillna("---").astype(str)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_ex.to_excel(writer, index=False, sheet_name='Escala')
                workbook = writer.book
                ws = writer.sheets['Escala']
                # Formatos
                f_header = workbook.add_format({'bold':True,'fg_color':'#1F4E78','font_color':'white','border':1,'align':'center'})
                f_qua = workbook.add_format({'bg_color':'#EBF1DE','border':1}) # Verde Claro
                f_sab = workbook.add_format({'bg_color':'#F2F2F2','border':1}) # Cinza
                f_dom = workbook.add_format({'bg_color':'#FFF2CC','border':1}) # Amarelo
                f_ceia = workbook.add_format({'bg_color':'#D9E1F2','border':1,'bold':True}) # Azul Ceia
                
                for col_num, val in enumerate(df_ex.columns.values): ws.write(0, col_num, val, f_header)
                
                for r_idx in range(len(df_ex)):
                    data_str = df_ex.iloc[r_idx, 0]
                    # Lógica de cor por dia
                    current_fmt = workbook.add_format({'border':1})
                    if "(Qua)" in data_str: current_fmt = f_qua
                    elif "(Sáb)" in data_str: current_fmt = f_sab
                    elif "(Dom)" in data_str: current_fmt = f_dom
                    
                    if data_ceia.strftime('%d/%m/%Y') in data_str: current_fmt = f_ceia
                    
                    for c_idx in range(len(df_ex.columns)):
                        ws.write(r_idx + 1, c_idx, df_ex.iloc[r_idx, c_idx], current_fmt)
                ws.set_column(0, 15, 20)
            st.download_button("📥 Excel Colorido por Dia", output.getvalue(), "escala_colorida.xlsx")

        with c_img:
            df_img = st.session_state.escala_gerada.fillna("---")
            # Aumento do tamanho da figura para evitar cortes
            fig, ax = plt.subplots(figsize=(18, len(df_img)*0.8 + 2)) 
            ax.axis('off')
            table = ax.table(cellText=df_img.values, colLabels=df_img.columns, loc='center', cellLoc='center', colColours=['#1F4E78']*len(df_img.columns))
            table.auto_set_font_size(False)
            table.set_fontsize(11)
            table.scale(1.2, 2.5) # Células bem altas para caber nomes
            
            for (i, j), cell in table.get_celld().items():
                if i == 0: cell.set_text_props(color='white', weight='bold')
                else:
                    d_text = df_img.iloc[i-1, 0]
                    if data_ceia.strftime('%d/%m/%Y') in d_text: cell.set_facecolor('#D9E1F2')
                    elif "(Qua)" in d_text: cell.set_facecolor('#EBF1DE')
                    elif "(Dom)" in d_text: cell.set_facecolor('#FFF2CC')
            
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', dpi=300) # Alta definição
            st.download_button("📸 Imagem p/ WhatsApp (HD)", buf.getvalue(), "escala_diaconato.png")

else: st.info("Suba o arquivo master para começar.")
