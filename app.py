import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala - Diaconato", layout="wide")

# Título em Português (Requisito 4)
st.title("⛪ Gerador de Escala de Diaconato")
st.markdown("---")

def obter_primeiro_domingo(ano, mes):
    d = date(ano, mes, 1)
    while d.weekday() != 6:
        d += timedelta(days=1)
    return d

# --- SIDEBAR / BARRA LATERAL ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")

if arquivo_carregado:
    df_membros = pd.read_csv(arquivo_carregado)
    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    st.sidebar.header("2. Configurações do Mês")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=2026)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=0, format_func=lambda x: [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_culto = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    sugestao_ceia = obter_primeiro_domingo(ano, mes)
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=sugestao_ceia)

    # --- PONTO 1: Impedir pessoas juntas ---
    st.sidebar.header("3. Restrições de Duplas")
    duplas_impedidas = st.sidebar.multiselect("Não escalar estas pessoas no MESMO dia:", nomes_membros)
    # Se selecionar 2 ou mais, o sistema garantirá que se um for escalado, o outro não será.

    # --- PONTO 2: Restrição de Função ---
    st.sidebar.header("4. Restrição de Função")
    membro_restrito = st.sidebar.selectbox("Membro com restrição:", ["Nenhum"] + nomes_membros)
    funcao_proibida = st.sidebar.multiselect("Não escalar este membro em:", 
                                            ["Portaria 1 (Rua)", "Portaria 2", "Frente Templo", "Abertura", "Santa Ceia"])

    # --- PONTO 3: Indisponibilidade Temporal ---
    st.sidebar.header("5. Indisponibilidade de Membros")
    membro_ausente = st.sidebar.selectbox("Membro ausente/férias:", ["Nenhum"] + nomes_membros)
    inicio_ausencia = st.sidebar.date_input("Início da ausência", value=date(ano, mes, 1))
    fim_ausencia = st.sidebar.date_input("Fim da ausência", value=date(ano, mes, 28))

    if st.sidebar.button("Gerar Escala"):
        inicio_mes = datetime(ano, mes, 1)
        if mes == 12: fim_mes = datetime(ano + 1, 1, 1) - pd.Timedelta(days=1)
        else: fim_mes = datetime(ano, mes + 1, 1) - pd.Timedelta(days=1)
            
        datas = pd.date_range(inicio_mes, fim_mes)
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        
        for data in datas:
            dia_semana_num = data.weekday()
            nome_coluna_dia = mapa_dias.get(dia_semana_num)
            
            if nome_coluna_dia in dias_culto:
                # 1. Filtro Geral de Disponibilidade (CSV + Ausência Temporal)
                candidatos_dia = df_membros[df_membros[nome_coluna_dia] != "NÃO"].copy()
                
                # Aplica Restrição do PONTO 3 (Ausência em datas específicas)
                if membro_ausente != "Nenhum":
                    if inicio_ausencia <= data.date() <= fim_ausencia:
                        candidatos_dia = candidatos_dia[candidatos_dia['Nome'] != membro_ausente]

                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                escalados_no_dia = {} 

                # Definindo Vagas
                if nome_coluna_dia == "Domingo":
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"]
                else:
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    # Filtra quem já foi escalado hoje
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    
                    # PONTO 1: Filtro de Duplas Impedidas
                    nomes_ja_escalados_hoje = list(escalados_no_dia.keys())
                    ja_tem_alguem_da_dupla = any(nome in duplas_impedidas for nome in nomes_ja_escalados_hoje)
                    if ja_tem_alguem_da_dupla:
                        candidatos = candidatos[~candidatos['Nome'].isin(duplas_impedidas)]

                    # PONTO 2: Filtro de Restrição de Função
                    if membro_restrito != "Nenhum":
                        # Simplifica a verificação de função proibida
                        for r_funcao in funcao_proibida:
                            if r_funcao in vaga:
                                candidatos = candidatos[candidatos['Nome'] != membro_restrito]

                    # Filtros de Sexo (Domingos)
                    if vaga == "Frente Templo (M)": candidatos = candidatos[candidatos['Sexo'] == 'M']
                    elif vaga == "Frente Templo (F)": candidatos = candidatos[candidatos['Sexo'] == 'F']
                    
                    # Ordenação por Equidade
                    candidatos = candidatos.sort_values(by='escalas_no_mes')

                    if not candidatos.empty:
                        escolhido = candidatos.iloc[0]
                        escalados_no_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # Santa Ceia e Abertura
                if data.date() == data_ceia:
                    aptos = [m for m in escalados_no_dia.keys() if m != dia_escala.get("Portaria 1 (Rua)")]
                    # Restrição de Função na Ceia
                    if membro_restrito != "Nenhum" and "Santa Ceia" in funcao_proibida:
                        aptos = [m for m in aptos if m != membro_restrito]
                        
                    homens = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    mulheres = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    dia_escala["Servir Santa Ceia"] = ", ".join(homens + mulheres)
                
                # Abertura
                c_abertura = candidatos_dia[(candidatos_dia['Abertura'] == "SIM") & (candidatos_dia['Nome'] != dia_escala.get("Portaria 1 (Rua)"))]
                # Se houver restrição de função para abertura
                if membro_restrito != "Nenhum" and "Abertura" in funcao_proibida:
                    c_abertura = c_abertura[c_abertura['Nome'] != membro_restrito]
                
                ja_no_templo = [n for n in escalados_no_dia.keys() if n in c_abertura['Nome'].values]
                dia_escala["Abertura"] = ja_no_templo[0] if ja_no_templo else (c_abertura[~c_abertura['Nome'].isin(escalados_no_dia.keys())].iloc[0]['Nome'] if not c_abertura[~c_abertura['Nome'].isin(escalados_no_dia.keys())].empty else "---")
                
                escala_final.append(dia_escala)

        df_resultado = pd.DataFrame(escala_final)
        st.subheader(f"Escala Gerada")
        st.dataframe(df_resultado, use_container_width=True)

        # Requisito 4: Arquivo Excel em Português
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Escala_Diaconato')
        
        st.download_button(label="📥 Baixar Escala (Excel)", data=output.getvalue(), 
                           file_name=f"escala_diaconato_{mes}_{ano}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Aguardando arquivo CSV para iniciar.")
