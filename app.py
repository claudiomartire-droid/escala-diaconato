import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Escala Diaconato V6.6", layout="wide")

if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None
if 'df_ausencias' not in st.session_state:
    # Inicializa com tipos definidos para evitar erro de compatibilidade
    st.session_state.df_ausencias = pd.DataFrame(columns=["Membro", "Início", "Fim"])

st.title("⛪ Gerador de Escala de Diaconato (Versão 6.6)")

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
    
    # Processamento de Ranking
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
    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia)

    # Captura de Regras
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

    st.subheader("📋 Conferência de Regras")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Ranking Ceia"])
    with t1: st.dataframe(pd.DataFrame(regras_duplas), use_container_width=True)
    with t2: st.dataframe(pd.DataFrame(regras_funcao), use_container_width=True)
    with t3: st.dataframe(df_membros[['Nome', 'historico_ceia']].sort_values('historico_ceia'), use_container_width=True)

    # --- 2. CONFIGURAÇÕES ---
    st.sidebar.header("2. Configurações")
    hoje = datetime.now()
    ano_sel = st.sidebar.number_input("Ano", 2025, 2030, hoje.year)
    mes_idx = st.sidebar.selectbox("Mês", range(1, 13), index=(hoje.month-1), format_func=lambda x: LISTA_MESES[x-1])
    nome_mes_sel = LISTA_MESES[mes_idx-1]
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano_sel, mes_idx))
    
    data_ini_mes = date(ano_sel, mes_idx, 1)
    data_fim_mes = (date(ano_sel + (1 if mes_idx==12 else 0), 1 if mes_idx==12 else mes_idx+1, 1) - timedelta(days=1))
    datas_excluir = st.sidebar.multiselect("Excluir Datas", options=pd.date_range(data_ini_mes, data_fim_mes), format_func=lambda x: x.strftime('%d/%m/%Y'))

    # --- 3. FÉRIAS / AUSÊNCIAS (FIXED V6.6) ---
    st.sidebar.header("3. Férias / Ausências")
    
    # Editor sem tentativa de mutação direta para evitar o erro de API
    edit_ausencias = st.sidebar.data_editor(
        st.session_state.df_ausencias,
        num_rows="dynamic",
        column_config={
            "Membro": st.column_config.SelectboxColumn("Membro", options=nomes_membros, required=True),
            "Início": st.column_config.DateColumn("Início", format="DD/MM/YYYY", required=True),
            "Fim": st.column_config.DateColumn("Fim", format="DD/MM/YYYY", required=False),
        },
        key="ausencias_editor_v6"
    )
    st.session_state.df_ausencias = edit_ausencias

    # --- MOTOR DE GERAÇÃO ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        datas_mes = pd.date_range(data_ini_mes, data_fim_mes)
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        ultima_escala = {nome: -10 for nome in nomes_membros} 
        membros_ultimo_culto = []

        for dia_idx, data in enumerate(datas_mes):
            data_atual = data.date()
            if any(data_atual == d.date() for d in datas_excluir): continue
            
            mapa = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
            nome_col_dia = mapa.get(data.weekday())

            if nome_col_dia in dias_semana:
                cands = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # Filtro Ausências com Lógica Inteligente (Início=Fim se vazio)
                for _, aus in st.session_state.df_ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']):
                        d_i = aus['Início']
                        # Se Fim estiver vazio, assume o próprio dia de Início
                        d_f = aus['Fim'] if pd.notna(aus['Fim']) else d_i
                        
                        # Normalização para date
                        if not isinstance(d_i, date): d_i = pd.to_datetime(d_i).date()
                        if not isinstance(d_f, date): d_f = pd.to_datetime(d_f).date()
                        
                        if d_i <= data_atual <= d_f:
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
                    for r in regras_duplas:
                        if r['Membro'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Evitar']]
                        if r['Evitar'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Membro']]
                    for rf in regras_funcao:
                        lista_res = [item.strip().lower() for item in rf['Restrição'].split(',')]
                        for termo in lista_res:
                            if termo != "" and termo in v.lower():
                                v_cands = v_cands[v_cands['Nome'] != rf['Membro']]
                                break

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
                for rf in regras_funcao:
                    lista_res = [item.strip().lower() for item in rf['Restrição'].split(',')]
                    if any(t in "abertura" for t in lista_res if t != ""):
                        aptos_ab = aptos_ab[aptos_ab['Nome'] != rf['Membro']]

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
                    dia_escala["Santa Ceia"] = "\n".join(escalados_dia[:4])
                
                escala_final.append(dia_escala)
                membros_ultimo_culto = escalados_dia

        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        st.session_state.df_memoria = df_membros[['Nome', 'historico_ceia']]

    # --- DOWNLOADS E RENDERIZAÇÃO ---
    if st.session_state.escala_gerada is not None:
        st.subheader(f"🗓️ Escala Gerada - {nome_mes_sel}")
        st.dataframe(st.session_state.escala_gerada, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            output = io.BytesIO()
            df_ex = st.session_state.escala_gerada.fillna("---").astype(str)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_ex.to_excel(writer, index=False, sheet_name='Escala')
                wb, ws = writer.book, writer.sheets['Escala']
                f_h = wb.add_format({'bold':True,'fg_color':'#1F4E78','font_color':'white','border':1,'align':'center'})
                f_q, f_s, f_d, f_c = wb.add_format({'bg_color':'#EBF1DE','border':1}), wb.add_format({'bg_color':'#F2F2F2','border':1}), wb.add_format({'bg_color':'#FFF2CC','border':1}), wb.add_format({'bg_color':'#D9E1F2','border':1,'bold':True})
                for col_num, val in enumerate(df_ex.columns.values): ws.write(0, col_num, val, f_h)
                for r_idx in range(len(df_ex)):
                    d_cell = df_ex.iloc[r_idx, 0]
                    fmt = wb.add_format({'border':1})
                    if data_ceia.strftime('%d/%m/%Y') in d_cell: fmt = f_c
                    elif "(Qua)" in d_cell: fmt = f_q
                    elif "(Sáb)" in d_cell: fmt = f_s
                    elif "(Dom)" in d_cell: fmt = f_d
                    for c_idx in range(len(df_ex.columns)): ws.write(r_idx+1, c_idx, df_ex.iloc[r_idx, c_idx], fmt)
                ws.set_column(0, 15, 25)
            st.download_button("📥 Excel Colorido", output.getvalue(), f"Escala_{nome_mes_sel}.xlsx")
        
        with c2:
            df_img = st.session_state.escala_gerada.fillna("---").copy()
            fig, ax = plt.subplots(figsize=(24, len(df_img) * 1.5 + 2))
            ax.axis('off')
            tab = ax.table(cellText=df_img.values, colLabels=df_img.columns, loc='center', cellLoc='center', colColours=['#1F4E78']*len(df_img.columns))
            tab.auto_set_font_size(False); tab.set_fontsize(11); tab.scale(1.2, 5.5)
            for (i, j), cell in tab.get_celld().items():
                if i == 0: cell.set_text_props(color='white', weight='bold')
                elif i > 0:
                    dt = df_img.iloc[i-1, 0]
                    if data_ceia.strftime('%d/%m/%Y') in dt: cell.set_facecolor('#D9E1F2')
                    elif "(Qua)" in dt: cell.set_facecolor('#EBF1DE')
                    elif "(Sáb)" in dt: cell.set_facecolor('#F2F2F2')
                    elif "(Dom)" in dt: cell.set_facecolor('#FFF2CC')
            buf = io.BytesIO(); plt.savefig(buf, format='png', bbox_inches='tight', dpi=300)
            st.download_button("📸 Imagem WhatsApp", buf.getvalue(), f"Escala_{nome_mes_sel}.png")
        
        with c3:
            out_h = io.BytesIO()
            st.session_state.df_memoria.to_csv(out_h, index=False)
            st.download_button("💾 Baixar Histórico", out_h.getvalue(), f"historico_{nome_mes_sel}.csv")
else: st.info("Suba o arquivo membros_master.csv para começar.")
