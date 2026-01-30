import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Escala Diaconato V5.8", layout="wide")

if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.8)")

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
arquivos_historicos = st.sidebar.file_uploader("Suba o historico_consolidado.csv", type=["csv", "xlsx"], accept_multiple_files=True)

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Processamento de Equidade Contínua
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

    # --- TABS DE CONFERÊNCIA ---
    st.subheader("📋 Conferência de Regras e Equidade")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Ranking Santa Ceia"])
    
    with t1:
        regras_d = []
        if 'Nao_Escalar_Com' in df_membros.columns:
            for _, r in df_membros[df_membros['Nao_Escalar_Com'].notna()].iterrows():
                regras_d.append({"Membro": r['Nome'], "Evitar": r['Nao_Escalar_Com']})
        st.dataframe(pd.DataFrame(regras_d), use_container_width=True)

    with t2:
        regras_f = []
        if 'Funcao_Restrita' in df_membros.columns:
            for _, r in df_membros[df_membros['Funcao_Restrita'].notna()].iterrows():
                regras_f.append({"Membro": r['Nome'], "Restrição": r['Funcao_Restrita']})
        st.dataframe(pd.DataFrame(regras_f), use_container_width=True)

    with t3:
        st.dataframe(df_membros[['Nome', 'historico_ceia']].sort_values('historico_ceia'), use_container_width=True)

    # --- 2. CONFIGURAÇÕES ---
    st.sidebar.header("2. Configurações")
    hoje = datetime.now()
    ano_sel = st.sidebar.number_input("Ano", 2025, 2030, hoje.year + (1 if hoje.month == 12 else 0))
    mes_idx = st.sidebar.selectbox("Mês", range(1, 13), index=(hoje.month % 12), format_func=lambda x: LISTA_MESES[x-1])
    nome_mes_sel = LISTA_MESES[mes_idx-1]
    
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano_sel, mes_idx))
    
    data_ini = date(ano_sel, mes_idx, 1)
    data_fim = (date(ano_sel + (1 if mes_idx==12 else 0), 1 if mes_idx==12 else mes_idx+1, 1) - timedelta(days=1))
    datas_excluir = st.sidebar.multiselect("Datas de Exclusão (Não gerar escala)", options=pd.date_range(data_ini, data_fim), format_func=lambda x: x.strftime('%d/%m/%Y'))

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

    # --- MOTOR DE GERAÇÃO ---
    if st.sidebar.button("Gerar Escala Atualizada"):
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

                # Vagas Padrão
                if data.weekday() == 6: # Domingo
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"]
                else:
                    vagas = ["Portaria 1 (Rua)", "Portaria 2", "Frente Templo"]

                for v in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia)]
                    if "M" in v or "Rua" in v: v_cands = v_cands[v_cands['Sexo'] == 'M']
                    if "(F)" in v: v_cands = v_cands[v_cands['Sexo'] == 'F']
                    
                    v_cands = v_cands.sort_values(['historico_ceia', 'escalas_no_mes'])
                    if not v_cands.empty:
                        escolhido = v_cands.iloc[0]['Nome']
                        dia_escala[v] = escolhido
                        escalados_dia.append(escolhido)
                        df_membros.loc[df_membros['Nome'] == escolhido, 'escalas_no_mes'] += 1

                # Função Abertura (Peso 0.5)
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

                # Santa Ceia / Ornamentação
                if data_atual == data_ceia:
                    dia_escala["Santa Ceia"] = ", ".join(escalados_dia[:4])
                    aptos_orn = cands[(cands['Ornamentacao'] == "SIM") & (~cands['Nome'].isin(escalados_dia))]
                    if not aptos_orn.empty:
                        dia_escala["Ornamentação"] = aptos_orn.iloc[0]['Nome']

                escala_final.append(dia_escala)
                membros_ultimo_culto = escalados_dia

        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        st.session_state.df_memoria = df_membros[['Nome', 'historico_ceia']]

    # --- ÁREA DE EXIBIÇÃO E EXPORTAÇÃO ---
    if st.session_state.escala_gerada is not None:
        st.subheader(f"🗓️ Escala Gerada - {nome_mes_sel} {ano_sel}")
        st.dataframe(st.session_state.escala_gerada, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        
        with c1:
            # Excel com Cores por Dia
            output = io.BytesIO()
            df_ex = st.session_state.escala_gerada.fillna("---").astype(str)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_ex.to_excel(writer, index=False, sheet_name='Escala')
                wb, ws = writer.book, writer.sheets['Escala']
                
                # Estilos
                fmt_h = wb.add_format({'bold':True,'fg_color':'#1F4E78','font_color':'white','border':1,'align':'center'})
                fmt_qua = wb.add_format({'bg_color':'#EBF1DE','border':1,'align':'center'})
                fmt_sab = wb.add_format({'bg_color':'#F2F2F2','border':1,'align':'center'})
                fmt_dom = wb.add_format({'bg_color':'#FFF2CC','border':1,'align':'center'})
                fmt_ceia = wb.add_format({'bg_color':'#D9E1F2','border':1,'bold':True,'align':'center'})
                
                for col_num, val in enumerate(df_ex.columns.values): ws.write(0, col_num, val, fmt_h)
                
                for r_idx in range(len(df_ex)):
                    data_cell = df_ex.iloc[r_idx, 0]
                    f = wb.add_format({'border':1,'align':'center'})
                    if "(Qua)" in data_cell: f = fmt_qua
                    elif "(Sáb)" in data_cell: f = fmt_sab
                    elif "(Dom)" in data_cell: f = fmt_dom
                    if data_ceia.strftime('%d/%m/%Y') in data_cell: f = fmt_ceia
                    
                    for c_idx in range(len(df_ex.columns)):
                        ws.write(r_idx+1, c_idx, df_ex.iloc[r_idx, c_idx], f)
                ws.set_column(0, 15, 22)
            
            st.download_button("📥 Baixar Escala Formatada (Excel)", output.getvalue(), 
                               f"Escala_Diaconato_{nome_mes_sel}_{ano_sel}.xlsx", key="ex_58")

        with c2:
            # Imagem HD
            df_img = st.session_state.escala_gerada.fillna("---")
            fig, ax = plt.subplots(figsize=(20, len(df_img)*0.8 + 2))
            ax.axis('off')
            table = ax.table(cellText=df_img.values, colLabels=df_img.columns, loc='center', cellLoc='center', colColours=['#1F4E78']*len(df_img.columns))
            table.auto_set_font_size(False)
            table.set_fontsize(11)
            table.scale(1.2, 2.8) # Altura generosa para não cortar nomes
            
            for (i, j), cell in table.get_celld().items():
                if i == 0: cell.set_text_props(color='white', weight='bold')
                else:
                    d_txt = df_img.iloc[i-1, 0]
                    if data_ceia.strftime('%d/%m/%Y') in d_txt: cell.set_facecolor('#D9E1F2')
                    elif "(Qua)" in d_txt: cell.set_facecolor('#EBF1DE')
                    elif "(Dom)" in d_txt: cell.set_facecolor('#FFF2CC')
            
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', dpi=300)
            st.download_button("📸 Baixar Imagem (WhatsApp)", buf.getvalue(), 
                               f"Escala_Imagem_{nome_mes_sel}_{ano_sel}.png", key="img_58")

        with c3:
            # Histórico Consolidado
            out_h = io.BytesIO()
            st.session_state.df_memoria.to_csv(out_h, index=False)
            st.download_button("💾 Baixar Histórico Consolidado", out_h.getvalue(), 
                               f"historico_consolidado_{nome_mes_sel}_{ano_sel}.csv", key="hist_58")

else:
    st.info("Aguardando upload do arquivo membros_master.csv para iniciar.")
