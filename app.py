import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

st.set_page_config(page_title="Gerador de Escala - Diaconato V2.2", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 2.2)")
st.markdown("---")

# Função para encontrar o primeiro domingo do mês
def get_first_sunday(year, month):
    d = date(year, month, 1)
    while d.weekday() != 6: # 6 é Domingo
        d += timedelta(days=1)
    return d

# 1. Upload e Tratamento de Dados
st.sidebar.header("1. Base de Dados")
uploaded_file = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")

if uploaded_file:
    df_membros = pd.read_csv(uploaded_file)
    
    st.sidebar.header("2. Configurações do Mês")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=2025)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=datetime.now().month-1, format_func=lambda x: [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_culto_selecionados = st.sidebar.multiselect("Dias de Culto Semanais", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    # PONTO 3: Sugerir o primeiro domingo do mês como padrão
    sugestao_ceia = get_first_sunday(ano, mes)
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=sugestao_ceia)
    
    todas_datas_mes = pd.date_range(start=datetime(ano, mes, 1), periods=31)
    opcoes_exclusao = [d.to_pydatetime() for d in todas_datas_mes if d.month == mes]
    
    datas_excluidas = st.sidebar.multiselect(
        "Datas para EXCLUIR (Não gerar escala):",
        options=opcoes_exclusao,
        format_func=lambda x: x.strftime('%d/%m/%Y')
    )
    datas_excluidas_set = {d.date() for d in datas_excluidas}

    if st.sidebar.button("Gerar Escala Atualizada"):
        inicio_mes = datetime(ano, mes, 1)
        if mes == 12: fim_mes = datetime(ano + 1, 1, 1) - pd.Timedelta(days=1)
        else: fim_mes = datetime(ano, mes + 1, 1) - pd.Timedelta(days=1)
            
        datas = pd.date_range(inicio_mes, fim_mes)
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        
        for data in datas:
            if data.date() in datas_excluidas_set:
                continue
                
            dia_semana_num = data.weekday()
            nome_coluna_dia = mapa_dias.get(dia_semana_num)
            
            if nome_coluna_dia in dias_culto_selecionados:
                disponiveis = df_membros[df_membros[nome_coluna_dia] != "NÃO"].copy()
                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                escalados_no_dia = {} # Armazena Nome -> Dados do Membro (Series)

                # Definindo as vagas do dia
                if nome_coluna_dia == "Domingo":
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"]
                else:
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    candidatos = disponiveis[~disponiveis['Nome'].isin(escalados_no_dia.keys())]
                    
                    # PONTO 1: Homem e Mulher na Frente do Templo (Domingo)
                    if vaga == "Frente Templo (M)":
                        candidatos = candidatos[candidatos['Sexo'] == 'M']
                    elif vaga == "Frente Templo (F)":
                        candidatos = candidatos[candidatos['Sexo'] == 'F']
                    elif "Portaria 1" in vaga:
                        candidatos = candidatos.sort_values(by=['Sexo', 'escalas_no_mes'], ascending=[False, True])
                    
                    candidatos = candidatos.sort_values(by='escalas_no_mes')

                    if not candidatos.empty:
                        escolhido = candidatos.iloc[0]
                        escalados_no_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # PONTO 2: Santa Ceia (2 Homens e 2 Mulheres dos já escalados)
                if data.date() == data_ceia:
                    nomes_escalados = list(escalados_no_dia.keys())
                    # Filtra quem não está na Portaria 1 (Rua)
                    aptos = [m for m in nomes_escalados if "Portaria 1" not in dia_escala.values() or m != dia_escala.get("Portaria 1 (Rua)")]
                    
                    homens_ceia = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    mulheres_ceia = [m for m in aptos if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    
                    servidores = homens_ceia + mulheres_ceia
                    dia_escala["Servir Santa Ceia"] = ", ".join(servidores)
                    
                    ornamentadores = disponiveis[disponiveis['Ornamentacao'] == "SIM"].sort_values(by='escalas_no_mes')
                    dia_escala["Ornamentação"] = ornamentadores.iloc[0]['Nome'] if not ornamentadores.empty else "---"

                # PONTO 3 (Abertura continua sem peso)
                candidatos_abertura = disponiveis[disponiveis['Abertura'] == "SIM"]
                ja_no_templo = [n for n in escalados_no_dia.keys() if n in candidatos_abertura['Nome'].values and "Portaria 1" not in dia_escala.values()]
                
                if ja_no_templo:
                    dia_escala["Abertura"] = ja_no_templo[0]
                else:
                    sobras = candidatos_abertura[~candidatos_abertura['Nome'].isin(escalados_no_dia.keys())]
                    dia_escala["Abertura"] = sobras.iloc[0]['Nome'] if not sobras.empty else "FALTA PESSOAL"
                
                escala_final.append(dia_escala)

        df_resultado = pd.DataFrame(escala_final)
        st.subheader(f"Escala Gerada")
        st.dataframe(df_resultado, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Escala')
        
        st.download_button(label="📥 Baixar Escala (Excel)", data=output.getvalue(), 
                           file_name=f"escala_{mes}_{ano}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Por favor, carregue o arquivo 'membros_master.csv' na barra lateral.")
