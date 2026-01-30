import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Escala Diaconato V5.9.1", layout="wide")

if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.9.1)")

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
        # Tenta ler com diferentes encodings e separadores
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Processamento de Ranking Ceia
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

    # --- PROCESSAMENTO DE REGRAS (CORRIGIDO) ---
    regras_duplas = []
    col_dupla = [c for c in df_membros.columns if 'Nao_Escalar_Com' in c]
    if col_dupla:
        c_n = col_dupla[0]
        for _, row in df_membros[df_membros[c_n].notna()].iterrows():
            if str(row[c_n]).strip() != "" and str(row[c_n]).lower() != 'nan':
                regras_duplas.append({"Membro": row['Nome'], "Evitar Escalar Com": row[c_n]})

    regras_funcao = []
    col_funcao = [c for c in df_membros.columns if 'Funcao_Restrita' in c]
    if col_funcao:
        c_f = col_funcao[0]
        for _, row in df_membros[df_membros[c_f].notna()].iterrows():
            if str(row[c_f]).strip() != "" and str(row[c_f]).lower() != 'nan':
                regras_funcao.append({"Membro": row['Nome'], "Restrição": row[c_f]})

    # --- TABS DE CONFERÊNCIA ---
    st.subheader("📋 Conferência de Regras e Equidade")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Ranking Santa Ceia"])
    
    with t1:
        if regras_duplas: st.dataframe(pd.DataFrame(regras_duplas), use_container_width=True)
        else: st.info("Nenhuma restrição de dupla encontrada no arquivo.")

    with t2:
        if regras_funcao: st.dataframe(pd.DataFrame(regras_funcao), use_container_width=True)
        else: st.info("Nenhuma restrição de função encontrada no arquivo.")

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
    datas_excluir = st.sidebar.multiselect("Datas de Exclusão", options=pd.date_range(data_ini, data_fim), format_func=lambda x: x.strftime('%d/%m/%Y'))

    # --- 3. FÉRIAS / AUSÊNCIAS ---
    st.sidebar.header("3. Férias / Ausências")
    ausencias = st.sidebar.data_editor(pd.DataFrame(columns=["Membro", "Início", "Fim"]), column_config={"Membro": st.column_config.SelectboxColumn(options=nomes_membros), "Início": st.column_config.DateColumn(format="DD/MM/YYYY"), "Fim": st.column_config.DateColumn(format="DD/MM/YYYY")}, num_rows="dynamic")

    # --- MOTOR DE GERAÇÃO V5.9.1 ---
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
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)] # Sem sequência
                
                # Filtro Ausências
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Início']) and pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                        cands = cands[cands['Nome'] != aus['Membro']]

                cands['dias_de_folga'] = cands['Nome'].map(ultima_escala).apply(lambda x: dia_idx - x)

                dias_pt = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
                dia_escala = {"Data": f"{data.strftime('%d/%m/%Y')} ({dias_pt[data.weekday()]})"}
                escalados_dia = []

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if data.weekday() == 6 else ["Portaria 1 (Rua)", "Portaria 2", "Frente Templo"]

                for v in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia)]
                    
                    # Filtro de Sexo
                    if "M" in v or "Rua" in v: v_cands = v_cands[v_cands['Sexo'] == 'M']
                    if "(F)" in v: v_cands = v_cands[v_cands['Sexo'] == 'F']
                    
                    # Filtro de Duplas Impeditivas
                    for r in regras_duplas:
                        if r['Membro'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Evitar Escalar Com']]
                        if r['Evitar Escalar Com'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Membro']]
                    
                    # Filtro de Restrição de Função
                    for rf in regras_funcao:
                        if rf['Restrição'] in v: v_cands = v_cands[v_cands['Nome'] != rf['Membro']]

                    v_cands = v_cands.sort_values(['escalas_no_mes', 'dias_de_folga', 'historico_ceia'], ascending=[True, False, True])
                    
                    if not v_cands.empty:
                        escolhido = v_cands.iloc[0]['Nome']
                        dia_escala[v] = escolhido
                        escalados_dia.append(escolhido)
                        df_membros.loc[df_membros['Nome'] == escolhido, 'escalas_no_mes'] += 1
                        ultima_escala[escolhido] = dia_idx

                # Abertura (Regra: Não pode ser Portaria 1)
                port1 = dia_escala.get("Portaria 1 (Rua)")
                aptos_ab = cands[(cands['Abertura'] == "SIM") & (cands['Nome'] != port1)].copy()
                
                ja_no_dia = [n for n in escalados_dia if n in aptos_ab['Nome'].values and n != port1]
                if ja_no_dia:
                    dia_escala["Abertura"] = ja_no_dia[0]
                else:
                    sobra_ab = aptos_ab[~aptos_ab['Nome'].isin(escalados_dia)]
                    if not sobra_ab.empty:
                        escolhido_ab = sobra_ab.sort_values(['escalas_no_mes', 'dias_de_folga']).iloc[0]['Nome']
                        dia_escala["Abertura"] = escolhido_ab
                        df_membros.loc[df_membros['Nome'] == escolhido_ab, 'escalas_no_mes'] += 0.5
                        ultima_escala[escolhido_ab] = dia_idx
                        escalados_dia.append(escolhido_ab)

                if data_atual == data_ceia:
                    dia_escala["Santa Ceia"] = ", ".join(escalados_dia[:4])
                
                escala_final.append(dia_escala)
                membros_ultimo_culto = escalados_dia

        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        st.session_state.df_memoria = df_membros[['Nome', 'historico_ceia']]

    # --- DOWNLOADS ---
    if st.session_state.escala_gerada is not None:
        st.subheader(f"🗓️ Escala Gerada - {nome_mes_sel} {ano_sel}")
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
            st.download_button("📥 Baixar Excel Formatado", output.getvalue(), f"Escala_Diaconato_{nome_mes_sel}_{ano_sel}.xlsx")
        with c2:
            df_img = st.session_state.escala_gerada.fillna("---")
            fig, ax = plt.subplots(figsize=(20, len(df_img)*0.8 + 2))
            ax.axis('off')
            table = ax.table(cellText=df_img.values, colLabels=df_img.columns, loc='center', cellLoc='center', colColours=['#1F4E78']*len(df_img.columns))
            table.auto_set_font_size(False); table.set_fontsize(11); table.scale(1.2, 2.8)
            for (i, j), cell in table.get_celld().items():
                if i == 0: cell.set_text_props(color='white', weight='bold')
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', dpi=300)
            st.download_button("📸 Baixar Imagem WhatsApp", buf.getvalue(), f"Escala_Imagem_{nome_mes_sel}_{ano_sel}.png")
        with c3:
            out_h = io.BytesIO()
            st.session_state.df_memoria.to_csv(out_h, index=False)
            st.download_button("💾 Baixar Histórico Consolidado", out_h.getvalue(), f"historico_consolidado_{nome_mes_sel}_{ano_sel}.csv")
else:
    st.info("Suba o arquivo membros_master.csv para começar.")
