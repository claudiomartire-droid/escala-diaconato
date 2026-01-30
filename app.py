import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V5.2", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.2)")
st.info("📅 Datas padronizadas para DD/MM/AAAA e regra de descanso aplicada à Abertura.")

# --- LÓGICA DE DATA PADRÃO (REGRA DA 1ª SEMANA) ---
hoje = datetime.now()
if hoje.day <= 7:
    mes_padrao = hoje.month
    ano_padrao = hoje.year
else:
    if hoje.month == 12:
        mes_padrao = 1
        ano_padrao = hoje.year + 1
    else:
        mes_padrao = hoje.month + 1
        ano_padrao = hoje.year

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
        try:
            arquivo_carregado.seek(0)
            df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')
        except Exception:
            arquivo_carregado.seek(0)
            df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='latin-1')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Processamento de Regras
    regras_duplas_csv = []
    if 'Nao_Escalar_Com' in df_membros.columns:
        for _, row in df_membros[df_membros['Nao_Escalar_Com'].notna()].iterrows():
            p_b = str(row['Nao_Escalar_Com']).strip()
            if p_b and p_b.lower() != 'nan':
                regras_duplas_csv.append({"Membro": row['Nome'], "Evitar Escalar Com": p_b})

    regras_funcao_csv = []
    if 'Funcao_Restrita' in df_membros.columns:
        for _, row in df_membros[df_membros['Funcao_Restrita'].notna()].iterrows():
            lista_funcoes = [f.strip() for f in str(row['Funcao_Restrita']).split(',')]
            for func in lista_funcoes:
                if func and func.lower() != 'nan':
                    regras_funcao_csv.append({"Membro": row['Nome'], "Função Proibida": func})

    st.subheader("📋 Conferência de Regras do CSV")
    tab1, tab2 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função"])
    with tab1:
        if regras_duplas_csv: st.dataframe(pd.DataFrame(regras_duplas_csv), use_container_width=True)
        else: st.info("Sem duplas impeditivas.")
    with tab2:
        if regras_funcao_csv: st.dataframe(pd.DataFrame(regras_funcao_csv), use_container_width=True)
        else: st.info("Sem restrições de função.")

    # --- 2. INTERFACE LATERAL ---
    st.sidebar.header("2. Configurações")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=ano_padrao)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=mes_padrao-1, format_func=lambda x: [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_culto = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    # AJUSTE: Formato de exibição da Santa Ceia
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes), format="DD/MM/YYYY")

    st.sidebar.header("3. Férias / Ausências")
    df_vazio_ausencias = pd.DataFrame(columns=["Membro", "Início", "Fim"])
    ausencias_editadas = st.sidebar.data_editor(
        df_vazio_ausencias,
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True), 
            "Início": st.column_config.DateColumn(required=True, format="DD/MM/YYYY"), 
            "Fim": st.column_config.DateColumn(required=True, format="DD/MM/YYYY")
        },
        num_rows="dynamic", key="editor_ausencias_v52"
    )

    # --- 4. MOTOR DE GERAÇÃO ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        inicio_mes = datetime(ano, mes, 1)
        if mes == 12: proximo = datetime(ano + 1, 1, 1)
        else: proximo = datetime(ano, mes + 1, 1)
        datas = pd.date_range(inicio_mes, proximo - timedelta(days=1))
        
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        # Lista de membros escalados no culto anterior para garantir o descanso
        membros_ultimo_culto = []

        for data in datas:
            data_atual = data.date()
            nome_col_dia = mapa_dias.get(data.weekday())
            
            if nome_col_dia in dias_culto:
                # REGRA DE SEQUÊNCIA: Remove quem trabalhou no último culto
                candidatos_dia = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                candidatos_dia = candidatos_dia[~candidatos_dia['Nome'].isin(membros_ultimo_culto)]

                # Filtro de Ausências
                for _, aus in ausencias_editadas.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        try:
                            if pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                                candidatos_dia = candidatos_dia[candidatos_dia['Nome'] != aus['Membro']]
                        except: continue

                dia_escala = {"Data": data.strftime('%d/%m/%Y (%a)')}
                escalados_no_dia = {} 

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if nome_col_dia == "Domingo" else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                # --- ESCALA DOS POSTOS ---
                for vaga in vagas:
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    if vaga == "Portaria 1 (Rua)": candidatos = candidatos[candidatos['Sexo'] == 'M']
                    
                    for regra in regras_duplas_csv:
                        if regra['Membro'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != regra['Evitar Escalar Com']]
                        if regra['Evitar Escalar Com'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != regra['Membro']]

                    for rest in regras_funcao_csv:
                        if rest['Membro'] in candidatos['Nome'].values and rest['Função Proibida'] in vaga:
                            candidatos = candidatos[candidatos['Nome'] != rest['Membro']]

                    if "Frente Templo (M)" in vaga: candidatos = candidatos[candidatos['Sexo'] == 'M']
                    elif "Frente Templo (F)" in vaga: candidatos = candidatos[candidatos['Sexo'] == 'F']
                    
                    candidatos = candidatos.sort_values(by='escalas_no_mes')
                    if not candidatos.empty:
                        escolhido = candidatos.iloc[0]
                        escalados_no_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # --- LÓGICA DE ABERTURA (COM REGRA DE SEQUÊNCIA) ---
                # Filtra aptos que NÃO trabalharam no culto anterior e NÃO estão escalados na rua hoje
                membros_aptos_ab = candidatos_dia[candidatos_dia['Abertura'] == "SIM"].copy()
                restritos_ab = [r['Membro'] for r in regras_funcao_csv if r['Função Proibida'] == "Abertura"]
                membros_aptos_ab = membros_aptos_ab[~membros_aptos_ab['Nome'].isin(restritos_ab)]
                
                # Prioriza quem já está no dia (exceto rua)
                ja_escalados_ab = [n for n in escalados_no_dia.keys() if n in membros_aptos_ab['Nome'].values and n != dia_escala.get("Portaria 1 (Rua)")]
                
                if ja_escalados_ab:
                    dia_escala["Abertura"] = ja_escalados_ab[0]
                else:
                    # Busca quem sobrou dos candidatos do dia (que já respeitam a folga do culto anterior)
                    sobra_ab = membros_aptos_ab[~membros_aptos_ab['Nome'].isin(escalados_no_dia.keys())]
                    if not sobra_ab.empty:
                        escolhido_ab = sobra_ab.sort_values(by='escalas_no_mes').iloc[0]
                        dia_escala["Abertura"] = escolhido_ab['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido_ab['Nome'], 'escalas_no_mes'] += 0.5 # Conta meia escala para abertura solo
                        escalados_no_dia[escolhido_ab['Nome']] = escolhido_ab
                    else:
                        dia_escala["Abertura"] = "---"

                # Santa Ceia
                if data_atual == data_ceia:
                    aptos_ceia = list(escalados_no_dia.keys())
                    restritos_ceia = [r['Membro'] for r in regras_funcao_csv if r['Função Proibida'] == "Santa Ceia"]
                    aptos_ceia = [m for m in aptos_ceia if m not in restritos_ceia]
                    serv_h = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    serv_f = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    total_ceia = serv_h + serv_f
                    if len(total_ceia) < 4:
                        extras = [m for m in aptos_ceia if m not in total_ceia]
                        total_ceia = (total_ceia + extras)[:4]
                    dia_escala["Servir Santa Ceia"] = ", ".join(total_ceia)
                
                # Atualiza quem trabalhou para a folga do próximo culto
                membros_ultimo_culto = list(escalados_no_dia.keys())
                escala_final.append(dia_escala)

        st.subheader("🗓️ Escala Gerada")
        df_final = pd.DataFrame(escala_final)
        st.dataframe(df_final, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Escala')
        st.download_button(label="📥 Baixar Escala em Excel", data=output.getvalue(), file_name=f"escala_{mes}_{ano}.xlsx")
else:
    st.info("Aguardando upload do arquivo membros_master.csv.")
