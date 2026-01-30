import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V5.6", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.6)")

# --- INICIALIZAÇÃO DO ESTADO ---
if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None

# --- LÓGICA DE DATA PADRÃO ---
hoje = datetime.now()
if hoje.day <= 7:
    mes_padrao, ano_padrao = hoje.month, hoje.year
else:
    if hoje.month == 12:
        mes_padrao, ano_padrao = 1, hoje.year + 1
    else:
        mes_padrao, ano_padrao = hoje.month + 1, hoje.year

def obter_primeiro_domingo(ano, mes):
    d = date(ano, mes, 1)
    while d.weekday() != 6: d += timedelta(days=1)
    return d

# --- 1. CARGA DE DADOS (SIDEBAR) ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")
arquivos_historicos = st.sidebar.file_uploader("Suba históricos ou consolidado", type=["csv", "xlsx"], accept_multiple_files=True)

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except Exception:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Processamento de Equidade Contínua
    contagem_ceia_historico = {nome: 0 for nome in nomes_membros}
    if arquivos_historicos:
        for arq in arquivos_historicos:
            try:
                df_h = pd.read_csv(arq) if arq.name.endswith('.csv') else pd.read_excel(arq)
                if 'historico_ceia' in df_h.columns and 'Nome' in df_h.columns:
                    for _, row in df_h.iterrows():
                        if row['Nome'] in contagem_ceia_historico:
                            contagem_ceia_historico[row['Nome']] += row['historico_ceia']
                else:
                    cols_alvo = [c for c in df_h.columns if any(x in c for x in ["Santa Ceia", "Ornamentação"])]
                    for col in cols_alvo:
                        for celula in df_h[col].dropna().astype(str):
                            for nome in nomes_membros:
                                if nome in celula: contagem_ceia_historico[nome] += 1
            except: continue

    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia_historico)

    # Regras de Duplas e Funções
    regras_duplas = []
    if 'Nao_Escalar_Com' in df_membros.columns:
        for _, row in df_membros[df_membros['Nao_Escalar_Com'].notna()].iterrows():
            if str(row['Nao_Escalar_Com']).strip().lower() != 'nan':
                regras_duplas.append({"Membro": row['Nome'], "Evitar Escalar Com": row['Nao_Escalar_Com']})

    regras_funcao = []
    if 'Funcao_Restrita' in df_membros.columns:
        for _, row in df_membros[df_membros['Funcao_Restrita'].notna()].iterrows():
            funcs = [f.strip() for f in str(row['Funcao_Restrita']).split(',')]
            for f in funcs:
                if f and f.lower() != 'nan':
                    regras_funcao.append({"Membro": row['Nome'], "Função Proibida": f})

    # TABS DE CONFERÊNCIA
    st.subheader("📋 Conferência de Regras e Equidade")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Ranking Santa Ceia"])
    with t1: st.dataframe(pd.DataFrame(regras_duplas), use_container_width=True)
    with t2: st.dataframe(pd.DataFrame(regras_funcao), use_container_width=True)
    with t3: st.dataframe(df_membros[['Nome', 'historico_ceia']].sort_values(by='historico_ceia'), use_container_width=True)

    # --- 2. CONFIGURAÇÕES (SIDEBAR) ---
    st.sidebar.header("2. Configurações")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=ano_padrao)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=mes_padrao-1, format_func=lambda x: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    data_inicio_mes = date(ano, mes, 1)
    data_fim_mes = (date(ano + (1 if mes==12 else 0), 1 if mes==12 else mes+1, 1) - timedelta(days=1))
    datas_excluir = st.sidebar.multiselect("Datas para EXCLUIR", options=pd.date_range(data_inicio_mes, data_fim_mes), format_func=lambda x: x.strftime('%d/%m/%Y'))
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes))

    # --- 3. FÉRIAS / AUSÊNCIAS (CALENDÁRIO CORRIGIDO) ---
    st.sidebar.header("3. Férias / Ausências")
    ausencias = st.sidebar.data_editor(
        pd.DataFrame(columns=["Membro", "Início", "Fim"]),
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True),
            "Início": st.column_config.DateColumn(format="DD/MM/YYYY", required=True),
            "Fim": st.column_config.DateColumn(format="DD/MM/YYYY", required=True)
        },
        num_rows="dynamic"
    )

    # --- 4. MOTOR DE GERAÇÃO ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        datas_mes = pd.date_range(data_inicio_mes, data_fim_mes)
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        membros_ultimo_culto = []

        for data in datas_mes:
            data_atual = data.date()
            if any(data_atual == d.date() for d in datas_excluir): continue
            
            nome_col_dia = mapa_dias.get(data.weekday())
            if nome_col_dia in dias_semana:
                cands = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # Filtro de Ausências
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        if pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                            cands = cands[cands['Nome'] != aus['Membro']]

                dia_escala = {"Data": data.strftime('%d/%m/%Y (%a)')}
                escalados_dia = {}
                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if data.weekday() == 6 else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                # Escala de Portarias/Frente
                for vaga in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia.keys())]
                    if "Portaria 1" in vaga or "(M)" in vaga: v_cands = v_cands[v_cands['Sexo'] == 'M']
                    if "(F)" in vaga: v_cands = v_cands[v_cands['Sexo'] == 'F']
                    
                    for r in regras_duplas:
                        if r['Membro'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Evitar Escalar Com']]
                        if r['Evitar Escalar Com'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Membro']]
                    for rest in regras_funcao:
                        if rest['Função Proibida'] in vaga: v_cands = v_cands[v_cands['Nome'] != rest['Membro']]

                    # Ordenação de Equidade
                    ordem = ['historico_ceia', 'escalas_no_mes'] if data_atual == data_ceia else ['escalas_no_mes']
                    v_cands = v_cands.sort_values(by=ordem)

                    if not v_cands.empty:
                        escolhido = v_cands.iloc[0]
                        dia_escala[vaga] = escolhido['Nome']
                        escalados_dia[escolhido['Nome']] = escolhido
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else: dia_escala[vaga] = "FALTA PESSOAL"

                # RESTAURADO: FUNÇÃO ABERTURA (Peso 0.5)
                aptos_ab = cands[cands['Abertura'] == "SIM"].copy()
                # Remove restritos de função
                aptos_ab = aptos_ab[~aptos_ab['Nome'].isin([r['Membro'] for r in regras_funcao if r['Função Proibida'] == "Abertura"])]
                
                # Se alguém já escalado no dia pode abrir (exceto Portaria 1), priorizamos para não gastar mais gente
                ja_no_dia = [n for n in escalados_dia.keys() if n in aptos_ab['Nome'].values and n != dia_escala.get("Portaria 1 (Rua)")]
                if ja_no_dia:
                    dia_escala["Abertura"] = ja_no_dia[0]
                else:
                    sobra_ab = aptos_ab[~aptos_ab['Nome'].isin(escalados_dia.keys())]
                    if not sobra_ab.empty:
                        escolhido_ab = sobra_ab.sort_values(by=['historico_ceia', 'escalas_no_mes']).iloc[0]
                        dia_escala["Abertura"] = escolhido_ab['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido_ab['Nome'], 'escalas_no_mes'] += 0.5
                        escalados_dia[escolhido_ab['Nome']] = escolhido_ab
                    else: dia_escala["Abertura"] = "---"

                # Santa Ceia e Ornamentação
                if data_atual == data_ceia:
                    aptos_orn = cands[(cands['Ornamentacao'] == "SIM") & (~cands['Nome'].isin(escalados_dia.keys()))]
                    esc_orn = aptos_orn.sort_values(by=['historico_ceia', 'escalas_no_mes']).head(2)
                    dia_escala["Ornamentação"] = ", ".join(esc_orn['Nome'].tolist())
                    for n in esc_orn['Nome']: df_membros.loc[df_membros['Nome'] == n, 'escalas_no_mes'] += 0.5
                    
                    esc_nomes = sorted(list(escalados_dia.keys()), key=lambda x: df_membros.loc[df_membros['Nome'] == x, 'historico_ceia'].values[0])
                    dia_escala["Servir Santa Ceia"] = ", ".join(esc_nomes[:4])

                escala_final.append(dia_escala)
                membros_ultimo_culto = list(escalados_dia.keys())
        
        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        
        # Consolidação do Memória
        df_mem = df_membros[['Nome', 'historico_ceia']].copy()
        for nome in nomes_membros:
            linha_ceia = st.session_state.escala_gerada[st.session_state.escala_gerada['Data'].str.contains(data_ceia.strftime('%d/%m/%Y'))]
            if not linha_ceia.empty and nome in str(linha_ceia.iloc[0].to_dict()):
                df_mem.loc[df_mem['Nome'] == nome, 'historico_ceia'] += 1
        st.session_state.df_memoria = df_mem

    # --- EXIBIÇÃO ---
    if st.session_state.escala_gerada is not None:
        st.subheader("🗓️ Escala Mensal Gerada")
        st.dataframe(st.session_state.escala_gerada, use_container_width=True)
        c1, c2 = st.columns(2)
        with c1:
            out_esc = io.BytesIO()
            st.session_state.escala_gerada.to_excel(out_esc, index=False)
            st.download_button("📥 Baixar Escala (Excel)", out_esc.getvalue(), f"escala_{mes}_{ano}.xlsx", key="d_esc")
        with c2:
            out_h = io.BytesIO()
            st.session_state.df_memoria.to_csv(out_h, index=False)
            st.download_button("💾 Baixar Histórico Acumulado", out_h.getvalue(), "historico_consolidado.csv", key="d_hist")
else:
    st.info("Aguardando arquivo membros_master.csv.")
