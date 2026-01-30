import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V5.3", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.3)")
st.info("🔄 Restaurado: Agora você pode excluir datas específicas (ex: sábados sem culto) no menu lateral.")

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

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except Exception:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Conferência de Regras
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
    t1, t2 = st.tabs(["👥 Duplas", "🚫 Funções"])
    with t1: st.dataframe(pd.DataFrame(regras_duplas), use_container_width=True) if regras_duplas else st.info("Sem duplas.")
    with t2: st.dataframe(pd.DataFrame(regras_funcao), use_container_width=True) if regras_funcao else st.info("Sem restrições.")

    # --- 2. CONFIGURAÇÕES ---
    st.sidebar.header("2. Configurações")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=ano_padrao)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=mes_padrao-1, format_func=lambda x: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    # --- NOVO/RESTAURADO: DATAS DE EXCLUSÃO ---
    datas_excluir = st.sidebar.multiselect(
        "Datas para EXCLUIR (Sem Culto)",
        options=pd.date_range(date(ano, mes, 1), (date(ano, mes, 1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)),
        format_func=lambda x: x.strftime('%d/%m/%Y')
    )
    
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes), format="DD/MM/YYYY")

    st.sidebar.header("3. Férias / Ausências")
    df_vazio = pd.DataFrame(columns=["Membro", "Início", "Fim"])
    ausencias = st.sidebar.data_editor(
        df_vazio,
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True), 
            "Início": st.column_config.DateColumn(required=True, format="DD/MM/YYYY"), 
            "Fim": st.column_config.DateColumn(required=True, format="DD/MM/YYYY")
        },
        num_rows="dynamic", key="editor_v53"
    )

    # --- 4. MOTOR ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        datas_mes = pd.date_range(date(ano, mes, 1), (date(ano, mes, 1) + timedelta(days=32)).replace(day=1) - timedelta(days=1))
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        membros_ultimo_culto = []

        for data in datas_mes:
            data_atual = data.date()
            # Pula se a data estiver na lista de exclusão
            if any(data_atual == d.date() for d in datas_excluir):
                continue
                
            nome_col_dia = mapa_dias.get(data.weekday())
            if nome_col_dia in dias_semana:
                # Regra de Descanso: Remove quem trabalhou no último culto
                candidatos = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                candidatos = candidatos[~candidatos['Nome'].isin(membros_ultimo_culto)]

                # Filtro de Ausências
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
                    
                    cand_vaga = cand_vaga.sort_values(by='escalas_no_mes')
                    if not cand_vaga.empty:
                        escolhido = cand_vaga.iloc[0]
                        escalados_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # Abertura (Respeitando Descanso)
                aptos_ab = candidatos[candidatos['Abertura'] == "SIM"].copy()
                restritos_ab = [r['Membro'] for r in regras_funcao if r['Função Proibida'] == "Abertura"]
                aptos_ab = aptos_ab[~aptos_ab['Nome'].isin(restritos_ab)]
                
                # Prioridade para quem já está no dia (exceto Rua)
                ja_no_dia_ab = [n for n in escalados_dia.keys() if n in aptos_ab['Nome'].values and n != dia_escala.get("Portaria 1 (Rua)")]
                if ja_no_dia_ab:
                    dia_escala["Abertura"] = ja_no_dia_ab[0]
                else:
                    sobra_ab = aptos_ab[~aptos_ab['Nome'].isin(escalados_dia.keys())]
                    if not sobra_ab.empty:
                        escolhido_ab = sobra_ab.sort_values(by='escalas_no_mes').iloc[0]
                        dia_escala["Abertura"] = escolhido_ab['Nome']
                        escalados_dia[escolhido_ab['Nome']] = escolhido_ab
                    else: dia_escala["Abertura"] = "---"

                # Santa Ceia
                if data_atual == data_ceia:
                    aptos_ceia = [m for m in escalados_dia.keys() if m not in [r['Membro'] for r in regras_funcao if r['Função Proibida'] == "Santa Ceia"]]
                    h = [m for m in aptos_ceia if escalados_dia[m]['Sexo'] == 'M'][:2]
                    f = [m for m in aptos_ceia if escalados_dia[m]['Sexo'] == 'F'][:2]
                    total = h + f
                    if len(total) < 4: total = (total + [m for m in aptos_ceia if m not in total])[:4]
                    dia_escala["Servir Santa Ceia"] = ", ".join(total)
                
                membros_ultimo_culto = list(escalados_dia.keys())
                escala_final.append(dia_escala)

        st.subheader("🗓️ Escala Gerada")
        df_res = pd.DataFrame(escala_final)
        st.dataframe(df_res, use_container_width=True)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as wr: df_res.to_excel(wr, index=False)
        st.download_button("📥 Baixar Excel", out.getvalue(), f"escala_{mes}_{ano}.xlsx")
else:
    st.info("Aguardando arquivo CSV.")
