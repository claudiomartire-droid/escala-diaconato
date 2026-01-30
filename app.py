import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V4.7", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 4.7)")
st.info("💡 Santa Ceia: Seleção automática de 2 homens e 2 mulheres entre os escalados do dia.")

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
    
    # Processamento de Regras do CSV
    regras_duplas_csv = []
    if 'Nao_Escalar_Com' in df_membros.columns:
        for _, row in df_membros[df_membros['Nao_Escalar_Com'].notna()].iterrows():
            if str(row['Nao_Escalar_Com']).strip() and str(row['Nao_Escalar_Com']).lower() != 'nan':
                regras_duplas_csv.append({"Pessoa A": row['Nome'], "Pessoa B": row['Nao_Escalar_Com']})

    regras_funcao_csv = []
    if 'Funcao_Restrita' in df_membros.columns:
        for _, row in df_membros[df_membros['Funcao_Restrita'].notna()].iterrows():
            lista_funcoes = [f.strip() for f in str(row['Funcao_Restrita']).split(',')]
            for func in lista_funcoes:
                if func and func.lower() != 'nan':
                    regras_funcao_csv.append({"Membro": row['Nome'], "Função Proibida": func})

    # --- 2. INTERFACE ---
    st.sidebar.header("2. Configurações")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=2026)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=0, format_func=lambda x: [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_culto = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes))

    # --- 3. TABELA DE FÉRIAS ---
    st.sidebar.header("3. Férias / Ausências")
    df_vazio_ausencias = pd.DataFrame(columns=["Membro", "Início", "Fim"])
    ausencias_editadas = st.sidebar.data_editor(
        df_vazio_ausencias,
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True), 
            "Início": st.column_config.DateColumn(required=True), 
            "Fim": st.column_config.DateColumn(required=True)
        },
        num_rows="dynamic", key="editor_ausencias_v47"
    )

    if st.sidebar.button("Gerar Escala Atualizada"):
        inicio_mes = datetime(ano, mes, 1)
        if mes == 12: proximo = datetime(ano + 1, 1, 1)
        else: proximo = datetime(ano, mes + 1, 1)
        datas = pd.date_range(inicio_mes, proximo - timedelta(days=1))
        
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        ultimos_escalados = []

        for data in datas:
            data_atual = data.date()
            nome_col_dia = mapa_dias.get(data.weekday())
            
            if nome_col_dia in dias_culto:
                candidatos_dia = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                candidatos_dia = candidatos_dia[~candidatos_dia['Nome'].isin(ultimos_escalados)]

                # Filtro de Ausências
                for _, aus in ausencias_editadas.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        try:
                            if pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                                candidatos_dia = candidatos_dia[candidatos_dia['Nome'] != aus['Membro']]
                        except: continue

                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                escalados_no_dia = {} 

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if nome_col_dia == "Domingo" else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                # --- ESCALA DOS POSTOS ---
                for vaga in vagas:
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    if vaga == "Portaria 1 (Rua)": candidatos = candidatos[candidatos['Sexo'] == 'M']
                    
                    for regra in regras_duplas_csv:
                        if regra['Pessoa A'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != regra['Pessoa B']]
                        if regra['Pessoa B'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != regra['Pessoa A']]

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

                # --- LÓGICA DE ABERTURA ---
                membros_aptos_abertura = candidatos_dia[candidatos_dia['Abertura'] == "SIM"].copy()
                restritos_abertura = [r['Membro'] for r in regras_funcao_csv if r['Função Proibida'] == "Abertura"]
                membros_aptos_abertura = membros_aptos_abertura[~membros_aptos_abertura['Nome'].isin(restritos_abertura)]
                ja_escalados_aptos = [n for n in escalados_no_dia.keys() if n in membros_aptos_abertura['Nome'].values and n != dia_escala.get("Portaria 1 (Rua)")]
                
                if ja_escalados_aptos:
                    dia_escala["Abertura"] = ja_escalados_aptos[0]
                else:
                    sobra_abertura = membros_aptos_abertura[~membros_aptos_abertura['Nome'].isin(escalados_no_dia.keys())]
                    dia_escala["Abertura"] = sobra_abertura.sort_values(by='escalas_no_mes').iloc[0]['Nome'] if not sobra_abertura.empty else "---"

                # --- LÓGICA DE SANTA CEIA (OTIMIZADA) ---
                if data_atual == data_ceia:
                    # Só pode servir quem já está escalado no dia e não está na Rua
                    aptos_ceia = [n for n in escalados_no_dia.keys() if n != dia_escala.get("Portaria 1 (Rua)")]
                    # Filtra quem tem restrição explícita de Santa Ceia no CSV
                    restritos_ceia = [r['Membro'] for r in regras_funcao_csv if r['Função Proibida'] == "Santa Ceia"]
                    aptos_ceia = [m for m in aptos_ceia if m not in restritos_ceia]
                    
                    # Tenta pegar 2 Homens e 2 Mulheres
                    servidores_h = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    servidores_f = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    
                    # Se faltar de um gênero, completa com o outro até ter 4 pessoas
                    total_servidores = servidores_h + servidores_f
                    if len(total_servidores) < 4:
                        extras = [m for m in aptos_ceia if m not in total_servidores]
                        total_servidores = (total_servidores + extras)[:4]
                    
                    dia_escala["Servir Santa Ceia"] = ", ".join(total_servidores)
                
                ultimos_escalados = list(escalados_no_dia.keys())
                escala_final.append(dia_escala)

        st.subheader(f"Escala Gerada")
        st.dataframe(pd.DataFrame(escala_final), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(escala_final).to_excel(writer, index=False, sheet_name='Escala')
        st.download_button(label="📥 Baixar Escala em Excel", data=output.getvalue(), file_name=f"escala_diaconato.xlsx")
else:
    st.info("Aguardando upload do arquivo membros_master.csv.")
