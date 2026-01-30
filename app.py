import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V4.3", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 4.3)")
st.info("💡 Carregamento automático: Regras e acentuação otimizadas para o seu arquivo CSV.")

def obter_primeiro_domingo(ano, mes):
    d = date(ano, mes, 1)
    while d.weekday() != 6: d += timedelta(days=1)
    return d

# --- 1. CARGA DE DADOS COM TRATAMENTO DE ERROS ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")

if arquivo_carregado:
    # Lógica para tratar o erro UnicodeDecodeError (Acentos do Excel/Windows)
    try:
        # Tenta a codificação padrão do Windows brasileiro (ISO-8859-1)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except Exception:
        try:
            # Segunda tentativa com UTF-8 (Padrão Web)
            arquivo_carregado.seek(0)
            df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')
        except Exception:
            # Terceira tentativa com Latin-1
            arquivo_carregado.seek(0)
            df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='latin-1')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # --- 2. PRÉ-PROCESSAMENTO DAS REGRAS DO CSV ---
    regras_duplas_csv = []
    if 'Nao_Escalar_Com' in df_membros.columns:
        for _, row in df_membros[df_membros['Nao_Escalar_Com'].notna()].iterrows():
            if str(row['Nao_Escalar_Com']).strip():
                regras_duplas_csv.append({"Pessoa A": row['Nome'], "Pessoa B": row['Nao_Escalar_Com']})

    regras_funcao_csv = []
    if 'Funcao_Restrita' in df_membros.columns:
        for _, row in df_membros[df_membros['Funcao_Restrita'].notna()].iterrows():
            lista_funcoes = [f.strip() for f in str(row['Funcao_Restrita']).split(',')]
            for func in lista_funcoes:
                if func and func.lower() != 'nan':
                    regras_funcao_csv.append({"Membro": row['Nome'], "Função Proibida": func})

    # --- 3. INTERFACE DE CONFIGURAÇÃO ---
    st.sidebar.header("2. Configurações do Mês")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=2026)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=0, format_func=lambda x: [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_culto = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes))

    st.sidebar.header("3. Ajustes de Regras")
    with st.sidebar.expander("Duplas (Lidas do CSV)"):
        df_duplas = st.data_editor(pd.DataFrame(regras_duplas_csv), num_rows="dynamic", key="ed_duplas")
    
    with st.sidebar.expander("Restrições de Função (Lidas do CSV)"):
        df_restricoes = st.data_editor(pd.DataFrame(regras_funcao_csv), num_rows="dynamic", key="ed_funcoes")

    st.sidebar.header("4. Férias / Ausências")
    df_ausencias = st.data_editor(pd.DataFrame(columns=["Membro", "Início", "Fim"]),
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros), 
            "Início": st.column_config.DateColumn(), 
            "Fim": st.column_config.DateColumn()
        },
        num_rows="dynamic", key="ed_ausencias")

    # --- 4. MOTOR DE GERAÇÃO DA ESCALA ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        inicio_mes = datetime(ano, mes, 1)
        if mes == 12: proximo = datetime(ano + 1, 1, 1)
        else: proximo = datetime(ano, mes + 1, 1)
        datas = pd.date_range(inicio_mes, proximo - timedelta(days=1))
        
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        ultimos_escalados = [] # Bloqueio de sequência de cultos

        for data in datas:
            data_atual = data.date()
            dia_semana_num = data.weekday()
            nome_col_dia = mapa_dias.get(dia_semana_num)
            
            if nome_col_dia in dias_culto:
                candidatos_dia = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                
                # Bloqueio de sequência (trabalhou no último culto, não trabalha hoje)
                candidatos_dia = candidatos_dia[~candidatos_dia['Nome'].isin(ultimos_escalados)]

                # Filtro de Ausências/Férias
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

                # Definindo postos por dia
                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if nome_col_dia == "Domingo" else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    
                    # Regra de Gênero: Apenas Homens na Rua
                    if vaga == "Portaria 1 (Rua)":
                        candidatos = candidatos[candidatos['Sexo'] == 'M']

                    # Regra de Duplas Impedidas
                    for _, dupla in df_duplas.iterrows():
                        if pd.notna(dupla.get('Pessoa A')) and pd.notna(dupla.get('Pessoa B')):
                            if dupla['Pessoa A'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa B']]
                            if dupla['Pessoa B'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa A']]

                    # Restrições de Função (CSV + Manual)
                    for _, rest in df_restricoes.iterrows():
                        if pd.notna(rest.get('Membro')) and pd.notna(rest.get('Função Proibida')) and rest['Função Proibida'] in vaga:
                            candidatos = candidatos[candidatos['Nome'] != rest['Membro']]

                    # Regras de Gênero: Frente Templo (Domingo)
                    if "Frente Templo (M)" in vaga: candidatos = candidatos[candidatos['Sexo'] == 'M']
                    elif "Frente Templo (F)" in vaga: candidatos = candidatos[candidatos['Sexo'] == 'F']
                    
                    # Equidade: prioriza quem tem menos escalas no mês
                    candidatos = candidatos.sort_values(by='escalas_no_mes')

                    if not candidatos.empty:
                        escolhido = candidatos.iloc[0]
                        escalados_no_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # Atualiza memória de sequência para o próximo culto
                ultimos_escalados = list(escalados_no_dia.keys())

                # --- 5. LÓGICA DE SANTA CEIA E ABERTURA ---
                if data_atual == data_ceia:
                    aptos = [m for m in escalados_no_dia.keys() if m != dia_escala.get("Portaria 1 (Rua)")]
                    for _, rest in df_restricoes.iterrows():
                        if rest['Função Proibida'] == "Santa Ceia":
                            aptos = [m for m in aptos if m != rest['Membro']]
                    h = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    f = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    dia_escala["Servir Santa Ceia"] = ", ".join(h + f)
                
                c_ab = candidatos_dia[(candidatos_dia['Abertura'] == "SIM") & (candidatos_dia['Nome'] != dia_escala.get("Portaria 1 (Rua)"))]
                for _, rest in df_restricoes.iterrows():
                    if rest['Função Proibida'] == "Abertura":
                        c_ab = c_ab[c_ab['Nome'] != rest['Membro']]
                
                ja_no_t = [n for n in escalados_no_dia.keys() if n in c_ab['Nome'].values]
                dia_escala["Abertura"] = ja_no_t[0] if ja_no_t else (c_ab[~c_ab['Nome'].isin(escalados_no_dia.keys())].iloc[0]['Nome'] if not c_ab[~c_ab['Nome'].isin(escalados_no_dia.keys())].empty else "---")
                
                escala_final.append(dia_escala)

        # --- 6. EXIBIÇÃO E EXCEL ---
        st.subheader(f"Escala Gerada")
        df_resultado = pd.DataFrame(escala_final)
        st.dataframe(df_resultado, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Escala')
        
        st.download_button(
            label="📥 Baixar Escala em Excel (PT-BR)",
            data=output.getvalue(),
            file_name=f"escala_diaconato_{mes}_{ano}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Aguardando upload do arquivo membros_master.csv para começar.")
