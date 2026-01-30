import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V5.6", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.6)")
st.info("⚖️ Pesos: Abertura e Ornamentação agora contam como 0.5 para o ranking de escalas.")

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

# --- 1. CARGA DE DADOS ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")

# NOVO: FUNCIONALIDADE SANTA CEIA - Upload de Histórico
arquivo_historico = st.sidebar.file_uploader("Opcional: Suba o histórico (Excel/CSV) dos últimos meses", type=["csv", "xlsx"])

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except Exception:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # NOVO: Processamento de Histórico para Equidade
    contagem_ceia_historico = {nome: 0 for nome in nomes_membros}
    if arquivo_historico:
        try:
            if arquivo_historico.name.endswith('.csv'):
                df_hist = pd.read_csv(arquivo_historico)
            else:
                df_hist = pd.read_excel(arquivo_historico)
            
            # Procura por colunas que contenham "Santa Ceia" ou "Ornamentação" no histórico
            colunas_ceia = [c for c in df_hist.columns if "Santa Ceia" in c or "Ornamentação" in c]
            for col in colunas_ceia:
                for lista_nomes in df_hist[col].dropna().astype(str):
                    for nome in nomes_membros:
                        if nome in lista_nomes:
                            contagem_ceia_historico[nome] += 1
            st.sidebar.success("✅ Histórico carregado para rodízio!")
        except Exception as e:
            st.sidebar.error(f"Erro ao ler histórico: {e}")

    # Adiciona o peso do histórico ao DataFrame principal
    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia_historico)

    # Conferência de Regras (Mantido original)
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

    st.subheader("📋 Conferência de Regras")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Equidade Santa Ceia"])
    with t1:
        if regras_duplas: st.dataframe(pd.DataFrame(regras_duplas), use_container_width=True)
        else: st.info("Sem duplas impeditivas.")
    with t2:
        if regras_funcao: st.dataframe(pd.DataFrame(regras_funcao), use_container_width=True)
        else: st.info("Sem restrições de função.")
    with t3:
        # Exibe para o usuário quem está na frente para o rodízio
        df_equidade = df_membros[['Nome', 'historico_ceia']].sort_values(by='historico_ceia')
        st.write("Membros com menos atuações em Santa Ceia (Prioridade de Escala):")
        st.dataframe(df_equidade, use_container_width=True)

    # --- 2. CONFIGURAÇÕES (Mantido original) ---
    st.sidebar.header("2. Configurações")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=ano_padrao)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=mes_padrao-1, format_func=lambda x: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    data_inicio_mes = date(ano, mes, 1)
    if mes == 12: data_fim_mes = date(ano + 1, 1, 1) - timedelta(days=1)
    else: data_fim_mes = date(ano, mes + 1, 1) - timedelta(days=1)
    
    datas_excluir = st.sidebar.multiselect("Datas para EXCLUIR", options=pd.date_range(data_inicio_mes, data_fim_mes), format_func=lambda x: x.strftime('%d/%m/%Y'))
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes), format="DD/MM/YYYY")

    st.sidebar.header("3. Férias / Ausências")
    ausencias = st.sidebar.data_editor(pd.DataFrame(columns=["Membro", "Início", "Fim"]), column_config={"Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True), "Início": st.column_config.DateColumn(required=True, format="DD/MM/YYYY"), "Fim": st.column_config.DateColumn(required=True, format="DD/MM/YYYY")}, num_rows="dynamic", key="editor_v56")

    # --- 4. MOTOR ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        datas_mes = pd.date_range(data_inicio_mes, data_fim_mes)
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        membros_ultimo_culto = []

        for data in datas_mes:
            data_atual = data.date()
            if any(data_atual == d.date() for d in datas_excluir): continue
                
            nome_col_dia = mapa_dias.get(data.weekday())
            if nome_col_dia in dias_semana:
                candidatos = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                candidatos = candidatos[~candidatos['Nome'].isin(membros_ultimo_culto)]

                # Filtro Ausências
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        if pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                            candidatos = candidatos[candidatos['Nome'] != aus['Membro']]

                dia_escala = {"Data": data.strftime('%d/%m/%Y (%a)')}
                escalados_dia = {} 

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if nome_col_dia == "Domingo" else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    cand_vaga = candidatos[~candidatos['Nome'].isin(escalados_dia.keys())]
                    if vaga == "Portaria 1 (Rua)": cand_vaga = cand_vaga[cand_vaga['Sexo'] == 'M']
                    
                    for r in regras_duplas:
                        if r['Membro'] in escalados_dia: cand_vaga = cand_vaga[cand_vaga['Nome'] != r['Evitar Escalar Com']]
                        if r['Evitar Escalar Com'] in escalados_dia: cand_vaga = cand_vaga[cand_vaga['Nome'] != r['Membro']]

                    for rest in regras_funcao:
                        if rest['Membro'] in cand_vaga['Nome'].values and rest['Função Proibida'] in vaga:
                            cand_vaga = cand_vaga[cand_vaga['Nome'] != rest['Membro']]

                    if "Frente Templo (M)" in vaga: cand_vaga = cand_vaga[cand_vaga['Sexo'] == 'M']
                    elif "Frente Templo (F)" in vaga: cand_vaga = cand_vaga[cand_vaga['Sexo'] == 'F']
                    
                    # NOVO: Se for domingo de ceia, o critério de desempate inclui o histórico de ceia
                    if data_atual == data_ceia:
                        cand_vaga = cand_vaga.sort_values(by=['historico_ceia', 'escalas_no_mes'])
                    else:
                        cand_vaga = cand_vaga.sort_values(by='escalas_no_mes')

                    if not cand_vaga.empty:
                        escolhido = cand_vaga.iloc[0]
                        escalados_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else: dia_escala[vaga] = "FALTA PESSOAL"

                # Abertura (Peso 0.5)
                aptos_ab = candidatos[candidatos['Abertura'] == "SIM"].copy()
                restritos_ab = [r['Membro'] for r in regras_funcao if r['Função Proibida'] == "Abertura"]
                aptos_ab = aptos_ab[~aptos_ab['Nome'].isin(restritos_ab)]
                
                ja_no_dia_ab = [n for n in escalados_dia.keys() if n in aptos_ab['Nome'].values and n != dia_escala.get("Portaria 1 (Rua)")]
                if ja_no_dia_ab:
                    dia_escala["Abertura"] = ja_no_dia_ab[0]
                else:
                    sobra_ab = aptos_ab[~aptos_ab['Nome'].isin(escalados_dia.keys())]
                    if not sobra_ab.empty:
                        # NOVO: Critério de rodízio na abertura também
                        sobra_ab = sobra_ab.sort_values(by=['historico_ceia', 'escalas_no_mes'])
                        escolhido_ab = sobra_ab.iloc[0]
                        dia_escala["Abertura"] = escolhido_ab['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido_ab['Nome'], 'escalas_no_mes'] += 0.5
                        escalados_dia[escolhido_ab['Nome']] = escolhido_ab
                    else: dia_escala["Abertura"] = "---"

                # --- SANTA CEIA E ORNAMENTAÇÃO ---
                if data_atual == data_ceia:
                    # 1. Ornamentação - Peso 0.5 + CRITÉRIO DE RODÍZIO HISTÓRICO
                    aptos_orn = candidatos[(candidatos['Ornamentacao'] == "SIM") & (~candidatos['Nome'].isin(escalados_dia.keys()))].copy()
                    if not aptos_orn.empty:
                        # NOVO: Aqui o 'historico_ceia' é o primeiro critério para garantir o rodízio
                        escolhidos_orn = aptos_orn.sort_values(by=['historico_ceia', 'escalas_no_mes']).head(2)
                        dia_escala["Ornamentação"] = ", ".join(escolhidos_orn['Nome'].tolist())
                        for n in escolhidos_orn['Nome']:
                            df_membros.loc[df_membros['Nome'] == n, 'escalas_no_mes'] += 0.5
                            escalados_dia[n] = n 
                    else:
                        dia_escala["Ornamentação"] = "FALTA PESSOAL"

                    # 2. Servir Santa Ceia (4 pessoas entre as escaladas no dia)
                    aptos_ceia = [m for m in escalados_dia.keys() if m not in [r['Membro'] for r in regras_funcao if r['Função Proibida'] == "Santa Ceia"]]
                    
                    # NOVO: Ordenar os aptos do dia pelo histórico para escolher quem servirá a ceia
                    def get_historico(nome): return df_membros.loc[df_membros['Nome'] == nome, 'historico_ceia'].values[0]
                    def get_sex(nome): return df_membros.loc[df_membros['Nome'] == nome, 'Sexo'].values[0]
                    
                    aptos_ceia_ordenados = sorted(aptos_ceia, key=lambda x: get_historico(x))
                    
                    h = [m for m in aptos_ceia_ordenados if get_sex(m) == 'M'][:2]
                    f = [m for m in aptos_ceia_ordenados if get_sex(m) == 'F'][:2]
                    total_ceia = (h + f)
                    if len(total_ceia) < 4: 
                        extras = [m for m in aptos_ceia_ordenados if m not in total_ceia]
                        total_ceia = (total_ceia + extras)[:4]
                    dia_escala["Servir Santa Ceia"] = ", ".join(total_ceia)
                
                membros_ultimo_culto = list(escalados_dia.keys())
                escala_final.append(dia_escala)

        st.subheader("🗓️ Escala Gerada")
        df_res = pd.DataFrame(escala_final)
        st.dataframe(df_res, use_container_width=True)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as wr: df_res.to_excel(wr, index=False)
        st.download_button("📥 Baixar Escala em Excel", out.getvalue(), f"escala_{mes}_{ano}.xlsx")
else:
    st.info("Aguardando arquivo membros_master.csv.")
