import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala - Diaconato Pro", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 3.0)")
st.markdown("---")

def obter_primeiro_domingo(ano, mes):
    d = date(ano, mes, 1)
    while d.weekday() != 6: d += timedelta(days=1)
    return d

# --- BARRA LATERAL ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")

if arquivo_carregado:
    df_membros = pd.read_csv(arquivo_carregado)
    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    st.sidebar.header("2. Configurações do Mês")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=2026)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=0, format_func=lambda x: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_culto = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes))

    # --- TABELAS DE REGRAS DINÂMICAS ---
    st.sidebar.header("3. Regras de Duplas")
    df_duplas = st.sidebar.data_editor(pd.DataFrame(columns=["Pessoa A", "Pessoa B"]),
        column_config={"Pessoa A": st.column_config.SelectboxColumn(options=nomes_membros), "Pessoa B": st.column_config.SelectboxColumn(options=nomes_membros)},
        num_rows="dynamic", key="ed_duplas")

    st.sidebar.header("4. Restrições de Função")
    df_restricoes = st.sidebar.data_editor(pd.DataFrame(columns=["Membro", "Função Proibida"]),
        column_config={"Membro": st.column_config.SelectboxColumn(options=nomes_membros), "Função Proibida": st.column_config.SelectboxColumn(options=["Portaria 1 (Rua)", "Portaria 2", "Frente Templo", "Abertura", "Santa Ceia"])},
        num_rows="dynamic", key="ed_funcoes")

    st.sidebar.header("5. Férias / Ausências")
    df_ausencias = st.sidebar.data_editor(pd.DataFrame(columns=["Membro", "Início", "Fim"]),
        column_config={"Membro": st.column_config.SelectboxColumn(options=nomes_membros), "Início": st.column_config.DateColumn(), "Fim": st.column_config.DateColumn()},
        num_rows="dynamic", key="ed_ausencias")

    if st.sidebar.button("Gerar Escala Atualizada"):
        # Define o período do mês
        inicio_foco = datetime(ano, mes, 1)
        if mes == 12: proximo_mes = datetime(ano + 1, 1, 1)
        else: proximo_mes = datetime(ano, mes + 1, 1)
        fim_foco = proximo_mes - timedelta(days=1)
        
        datas = pd.date_range(inicio_foco, fim_foco)
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        ultimos_escalados = [] # Lista para evitar sequência (Ponto 1)

        for data in datas:
            data_atual = data.date()
            dia_semana_num = data.weekday()
            nome_coluna_dia = mapa_dias.get(dia_semana_num)
            
            if nome_coluna_dia in dias_culto:
                # 1. Filtro base por disponibilidade no CSV
                candidatos_dia = df_membros[df_membros[nome_coluna_dia] != "NÃO"].copy()
                
                # --- APLICAÇÃO DA REGRA: NÃO REPETIR SEQUÊNCIA ---
                candidatos_dia = candidatos_dia[~candidatos_dia['Nome'].isin(ultimos_escalados)]

                # 2. Filtro de Ausências/Férias
                for _, aus in df_ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        try:
                            d_ini = pd.to_datetime(aus['Início']).date()
                            d_fim = pd.to_datetime(aus['Fim']).date()
                            if d_ini <= data_atual <= d_fim:
                                candidatos_dia = candidatos_dia[candidatos_dia['Nome'] != aus['Membro']]
                        except: continue

                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                escalados_no_dia = {} 

                # Define postos do dia
                if nome_coluna_dia == "Domingo":
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"]
                else:
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    
                    # --- APLICAÇÃO DA REGRA: APENAS HOMENS NA RUA ---
                    if vaga == "Portaria 1 (Rua)":
                        candidatos = candidatos[candidatos['Sexo'] == 'M']

                    # Regras de Duplas
                    for _, dupla in df_duplas.iterrows():
                        if pd.notna(dupla['Pessoa A']) and pd.notna(dupla['Pessoa B']):
                            if dupla['Pessoa A'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa B']]
                            if dupla['Pessoa B'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa A']]

                    # Restrições de Função individuais
                    for _, rest in df_restricoes.iterrows():
                        if pd.notna(rest['Membro']) and pd.notna(rest['Função Proibida']) and rest['Função Proibida'] in vaga:
                            candidatos = candidatos[candidatos['Nome'] != rest['Membro']]

                    # Gênero na Frente do Templo (Domingos)
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

                # Guardar quem trabalhou hoje para a próxima rodada
                ultimos_escalados = list(escalados_no_dia.keys())

                # Santa Ceia e Abertura
                if data_atual == data_ceia:
                    aptos = [m for m in escalados_no_dia.keys() if m != dia_escala.get("Portaria 1 (Rua)")]
                    h = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    f = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    dia_escala["Servir Santa Ceia"] = ", ".join(h + f)
                
                c_ab = candidatos_dia[(candidatos_dia['Abertura'] == "SIM") & (candidatos_dia['Nome'] != dia_escala.get("Portaria 1 (Rua)"))]
                ja_no_t = [n for n in escalados_no_dia.keys() if n in c_ab['Nome'].values]
                dia_escala["Abertura"] = ja_no_t[0] if ja_no_t else (c_ab[~c_ab['Nome'].isin(escalados_no_dia.keys())].iloc[0]['Nome'] if not c_ab[~c_ab['Nome'].isin(escalados_no_dia.keys())].empty else "---")
                
                escala_final.append(dia_escala)

        # Exibição e Download
        df_res = pd.DataFrame(escala_final)
        st.subheader(f"Escala Gerada")
        st.dataframe(df_res, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, index=False, sheet_name='Escala')
        st.download_button(label="📥 Baixar em Excel (PT-BR)", data=output.getvalue(), file_name=f"escala_diaconato_{mes}_{ano}.xlsx")
else:
    st.info("Aguardando o arquivo membros_master.csv para começar.")
