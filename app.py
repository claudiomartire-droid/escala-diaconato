import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

st.set_page_config(page_title="Gerador de Escala - Diaconato Pro", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Múltiplas Regras)")
st.markdown("---")

def obter_primeiro_domingo(ano, mes):
    d = date(ano, mes, 1)
    while d.weekday() != 6: d += timedelta(days=1)
    return d

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

    # --- PONTO 1: Múltiplas Duplas Impedidas ---
    st.sidebar.header("3. Regras de Duplas (Conflito)")
    st.sidebar.info("Adicione irmãos que NÃO podem ser escalados no mesmo dia (ex: Casais com filhos).")
    df_duplas = st.sidebar.data_editor(
        pd.DataFrame(columns=["Pessoa A", "Pessoa B"]),
        column_config={
            "Pessoa A": st.column_config.SelectboxColumn(options=nomes_membros, required=True),
            "Pessoa B": st.column_config.SelectboxColumn(options=nomes_membros, required=True),
        },
        num_rows="dynamic",
        key="editor_duplas"
    )

    # --- PONTO 2: Múltiplas Restrições de Função ---
    st.sidebar.header("4. Restrições de Função")
    df_restricoes = st.sidebar.data_editor(
        pd.DataFrame(columns=["Membro", "Função Proibida"]),
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True),
            "Função Proibida": st.column_config.SelectboxColumn(options=["Portaria 1 (Rua)", "Portaria 2", "Frente Templo", "Abertura", "Santa Ceia"], required=True),
        },
        num_rows="dynamic",
        key="editor_funcoes"
    )

    # --- PONTO 3: Múltiplas Ausências ---
    st.sidebar.header("5. Férias / Ausências")
    df_ausencias = st.sidebar.data_editor(
        pd.DataFrame(columns=["Membro", "Início", "Fim"]),
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True),
            "Início": st.column_config.DateColumn(required=True),
            "Fim": st.column_config.DateColumn(required=True),
        },
        num_rows="dynamic",
        key="editor_ausencias"
    )

    if st.sidebar.button("Gerar Escala"):
        datas = pd.date_range(datetime(ano, mes, 1), (datetime(ano, mes, 1) + timedelta(days=32)).replace(day=1) - timedelta(days=1))
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        
        for data in datas:
            dia_semana_num = data.weekday()
            nome_coluna_dia = mapa_dias.get(dia_semana_num)
            
            if nome_coluna_dia in dias_culto:
                # Filtro de Ausência (Multi-membros)
                candidatos_dia = df_membros[df_membros[nome_coluna_dia] != "NÃO"].copy()
                for _, aus in df_ausencias.iterrows():
                    if pd.notna(aus['Início']) and aus['Início'] <= data.date() <= aus['Fim']:
                        candidatos_dia = candidatos_dia[candidatos_dia['Nome'] != aus['Membro']]

                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                escalados_no_dia = {} 

                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if nome_coluna_dia == "Domingo" else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    
                    # Regra de Duplas
                    for _, dupla in df_duplas.iterrows():
                        if dupla['Pessoa A'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa B']]
                        if dupla['Pessoa B'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa A']]

                    # Regra de Restrição de Função
                    for _, rest in df_restricoes.iterrows():
                        if rest['Função Proibida'] in vaga:
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

                # Santa Ceia e Abertura com restrições
                if data.date() == data_ceia:
                    aptos_ceia = [m for m in escalados_no_dia.keys() if m != dia_escala.get("Portaria 1 (Rua)")]
                    for _, rest in df_restricoes.iterrows():
                        if rest['Função Proibida'] == "Santa Ceia": aptos_ceia = [m for m in aptos_ceia if m != rest['Membro']]
                    
                    h = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    f = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    dia_escala["Servir Santa Ceia"] = ", ".join(h + f)
                
                c_abertura = candidatos_dia[(candidatos_dia['Abertura'] == "SIM") & (candidatos_dia['Nome'] != dia_escala.get("Portaria 1 (Rua)"))]
                for _, rest in df_restricoes.iterrows():
                    if rest['Função Proibida'] == "Abertura": c_abertura = c_abertura[c_abertura['Nome'] != rest['Membro']]
                
                ja_no_templo = [n for n in escalados_no_dia.keys() if n in c_abertura['Nome'].values]
                dia_escala["Abertura"] = ja_no_templo[0] if ja_no_templo else (c_abertura[~c_abertura['Nome'].isin(escalados_no_dia.keys())].iloc[0]['Nome'] if not c_abertura[~c_abertura['Nome'].isin(escalados_no_dia.keys())].empty else "---")
                
                escala_final.append(dia_escala)

        df_resultado = pd.DataFrame(escala_final)
        st.subheader(f"Escala Gerada")
        st.dataframe(df_resultado, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Escala')
        st.download_button(label="📥 Baixar Escala (Excel)", data=output.getvalue(), file_name=f"escala_{mes}_{ano}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Aguardando arquivo CSV para iniciar.")
