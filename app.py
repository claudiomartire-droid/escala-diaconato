import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Escala Diaconato V5.9.2", layout="wide")

if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.9.2)")

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

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    # Limpeza básica de nomes para evitar erros de busca
    df_membros['Nome'] = df_membros['Nome'].astype(str).str.strip()
    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Ranking Ceia
    contagem_ceia = {nome: 0 for nome in nomes_membros}
    if arquivos_historicos:
        for arq in arquivos_historicos:
            try:
                df_h = pd.read_csv(arq) if arq.name.endswith('.csv') else pd.read_excel(arq)
                if 'historico_ceia' in df_h.columns:
                    for _, r in df_h.iterrows():
                        n = str(r['Nome']).strip()
                        if n in contagem_ceia: contagem_ceia[n] += r['historico_ceia']
            except: continue
    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia)

    # --- PROCESSAMENTO DE REGRAS ---
    regras_duplas = []
    col_dupla = [c for c in df_membros.columns if 'Nao_Escalar_Com' in c]
    if col_dupla:
        for _, row in df_membros[df_membros[col_dupla[0]].notna()].iterrows():
            m_alvo = str(row[col_dupla[0]]).strip()
            if m_alvo and m_alvo.lower() != 'nan':
                regras_duplas.append({"Membro": row['Nome'], "Evitar": m_alvo})

    regras_funcao = []
    col_funcao = [c for c in df_membros.columns if 'Funcao_Restrita' in c]
    if col_funcao:
        for _, row in df_membros[df_membros[col_funcao[0]].notna()].iterrows():
            restr = str(row[col_funcao[0]]).strip()
            if restr and restr.lower() != 'nan':
                regras_funcao.append({"Membro": row['Nome'], "Restrição": restr})

    # Tabs de Conferência
    st.subheader("📋 Conferência de Regras e Equidade")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Ranking Santa Ceia"])
    with t1: st.dataframe(pd.DataFrame(regras_duplas) if regras_duplas else pd.DataFrame(columns=["Info"]), use_container_width=True)
    with t2: st.dataframe(pd.DataFrame(regras_funcao) if regras_funcao else pd.DataFrame(columns=["Info"]), use_container_width=True)
    with t3: st.dataframe(df_membros[['Nome', 'historico_ceia']].sort_values('historico_ceia'), use_container_width=True)

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
    datas_excluir = st.sidebar.multiselect("Excluir Datas", options=pd.date_range(data_ini, data_fim), format_func=lambda x: x.strftime('%d/%m/%Y'))

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

    # --- MOTOR DE GERAÇÃO V5.9.2 ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        datas_mes = pd.date_range(data_ini, data_fim)
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        ultima_escala = {nome: -10 for nome in nomes_membros} 
        membros_ultimo_culto = []

        for dia_idx, data in enumerate(datas_mes):
            data_atual = data.date()
            if any(data_atual == d.date() for d in datas_excluir): continue
            
            mapa = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
            nome_col = mapa.get(data.weekday())

            if nome_col in dias_semana:
                cands = df_membros[df_membros[nome_col] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # FILTRO DE AUSÊNCIAS (CORRIGIDO)
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        d_ini = aus['Início'] if isinstance(aus['Início'], date) else pd.to_datetime(aus['Início']).date()
                        d_fim = aus['Fim'] if isinstance(aus['Fim'], date) else pd.to_datetime(aus['Fim']).date()
                        if d_ini <= data_atual <= d_fim:
                            cands = cands[cands['Nome'] != aus['Membro']]

                cands['folga'] = cands['Nome'].map(ultima_escala).apply(lambda x: dia_idx - x)
                dia_pt = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
                dia_escala = {"Data": f"{data.strftime('%d/%m/%Y')} ({dia_pt[data.weekday()]})"}
                escalados_dia = []

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if data.weekday() == 6 else ["Portaria 1 (Rua)", "Portaria 2", "Frente Templo"]

                for v in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia)]
                    if "M" in v or "Rua" in v: v_cands = v_cands[v_cands['Sexo'] == 'M']
                    if "(F)" in v: v_cands = v_cands[v_cands['Sexo'] == 'F']
                    
                    # Aplicar Regras de Duplas e Funções
                    for r in regras_duplas:
                        if r['Membro'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Evitar']]
                        if r['Evitar'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Membro']]
                    for rf in regras_funcao:
                        if rf['Restrição'] in v: v_cands = v_cands[v_cands['Nome'] != rf['Membro']]

                    v_cands = v_cands.sort_values(['escalas_no_mes', 'folga', 'historico_ceia'], ascending=[True, False, True])
                    if not v_cands.empty:
                        esc = v_cands.iloc[0]['Nome']
                        dia_escala[v] = esc
                        escalados_dia.append(esc)
                        df_membros.loc[df_membros['Nome'] == esc, 'escalas_no_mes'] += 1
                        ultima_escala[esc] = dia_idx

                # Abertura
                p1 = dia_escala.get("Portaria 1 (Rua)")
                aptos_ab = cands[(cands['Abertura'] == "SIM") & (cands['Nome'] != p1)].copy()
                ja_no_dia = [n for n in escalados_dia if n in aptos_ab['Nome'].values and n != p1]
                if ja_no_dia:
                    dia_escala["Abertura"] = ja_no_dia[0]
                else:
                    sobra_ab = aptos_ab[~aptos_ab['Nome'].isin(escalados_dia)]
                    if not sobra_ab.empty:
                        esc_ab = sobra_ab.sort_values(['escalas_no_mes', 'folga']).iloc[0]['Nome']
                        dia_escala["Abertura"] = esc_ab
                        df_membros.loc[df_membros['Nome'] == esc_ab, 'escalas_no_mes'] += 0.5
                        ultima_escala[esc_ab] = dia_idx
                        escalados_dia.append(esc_ab)

                if data_atual == data_ceia:
                    dia_escala["Santa Ceia"] = ", ".join(escalados_dia[:4])
                
                escala_final.append(dia_escala)
                membros_ultimo_culto = escalados_dia

        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        st.session_state.df_memoria = df_membros[['Nome', 'historico_ceia']]

    # --- DOWNLOADS ---
    if st.session_state.escala_gerada is not None:
        st.subheader(f"🗓️ Escala - {nome_mes_sel} {ano_sel}")
        st.dataframe(st.session_state.escala_gerada, use_container_width=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            output = io.BytesIO()
            df_ex = st.session_state.escala_gerada.fillna("---").astype(str)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_ex.to_excel(writer, index=False, sheet_name='Escala')
                wb, ws = writer.book, writer.sheets['Escala']
                f_h = wb.add_format({'bold':True,'fg_color':'#1F4E78','font_color':'white','border':1,'align':'center'})
                for col_num, val in enumerate(df_ex.columns.values): ws.write(0, col_num, val, f_h)
                ws.set_column(0, 15, 22)
            st.download_button("📥 Excel Formatado", output.getvalue(), f"Escala_{nome_mes_sel}_{ano_sel}.xlsx")
        with c2:
            fig, ax = plt.subplots(figsize=(20, len(st.session_state.escala_gerada)*0.8 + 2))
            ax.axis('off')
            tab = ax.table(cellText=st.session_state.escala_gerada.fillna("---").values, colLabels=st.session_state.escala_gerada.columns, loc='center', cellLoc='center', colColours=['#1F4E78']*len(st.session_state.escala_gerada.columns))
            tab.auto_set_font_size(False); tab.set_fontsize(11); tab.scale(1.2, 2.8)
            for (i, j), cell in tab.get_celld().items():
                if i == 0: cell.set_text_props(color='white', weight='bold')
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', dpi=300)
            st.download_button("📸 Imagem WhatsApp", buf.getvalue(), f"Escala_{nome_mes_sel}_{ano_sel}.png")
        with c3:
            out_h = io.BytesIO()
            st.session_state.df_memoria.to_csv(out_h, index=False)
            st.download_button("💾 Baixar Histórico", out_h.getvalue(), f"historico_{nome_mes_sel}_{ano_sel}.csv")
else: st.info("Suba o arquivo master.")
