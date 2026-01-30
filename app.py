import streamlit as st
import pandas as pd
import pandas.tseries.offsets as offsets
from datetime import datetime
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala - Diaconato", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato")
st.markdown("---")

# 1. Upload e Tratamento de Dados
st.sidebar.header("1. Base de Dados")
uploaded_file = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")

if uploaded_file:
    df_membros = pd.read_csv(uploaded_file)
    
    # Sidebar - Configurações de Data
    st.sidebar.header("2. Configurações do Mês")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=2025)
    mes = st.sidebar.selectbox("Mês", range(1, 13), format_func=lambda x: [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    dias_culto = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    data_ceia = st.sidebar.date_input("Data da Santa Ceia")

    if st.sidebar.button("Gerar Escala"):
        # Lógica para pegar todos os dias do mês
        inicio_mes = datetime(ano, mes, 1)
        if mes == 12:
            fim_mes = datetime(ano + 1, 1, 1) - pd.Timedelta(days=1)
        else:
            fim_mes = datetime(ano, mes + 1, 1) - pd.Timedelta(days=1)
            
        datas = pd.date_range(inicio_mes, fim_mes)
        
        # Mapeamento de nomes de dias (Ajustado para o CSV)
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        
        escala_final = []
        # Inicializa contador de escalas para equidade
        df_membros['escalas_no_mes'] = 0
        
        for data in datas:
            dia_semana_num = data.weekday() # 2=Quarta, 5=Sábado, 6=Domingo
            nome_coluna_dia = mapa_dias.get(dia_semana_num)
            
            if nome_coluna_dia in dias_culto:
                # Filtrar apenas disponíveis para este dia específico
                disponiveis = df_membros[df_membros[nome_coluna_dia] != "NÃO"].copy()
                
                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                
                # Funções a preencher
                funcoes = ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo", "Abertura"]
                if data.date() == data_ceia:
                    funcoes += ["Servir Santa Ceia", "Ornamentar Mesa"]

                escalados_do_dia = []

                for funcao in funcoes:
                    # Filtra quem ainda não foi escalado hoje
                    candidatos = disponiveis[~disponiveis['Nome'].isin(escalados_do_dia)]
                    
                    # Regras específicas
                    if funcao == "Portaria 1 (Rua)":
                        candidatos = candidatos.sort_values(by=['Sexo', 'escalas_no_mes'], ascending=[False, True]) # Prioriza 'M'
                    elif funcao == "Abertura":
                        candidatos = candidatos[candidatos['Abertura'] == "SIM"].sort_values(by='escalas_no_mes')
                    elif funcao == "Ornamentar Mesa":
                        candidatos = candidatos[candidatos['Ornamentacao'] == "SIM"].sort_values(by='escalas_no_mes')
                    else:
                        candidatos = candidatos.sort_values(by='escalas_no_mes')

                    if not candidatos.empty:
                        escolhido = candidatos.iloc[0]['Nome']
                        dia_escala[funcao] = escolhido
                        escalados_do_dia.append(escolhido)
                        # Atualiza contador no DF principal
                        df_membros.loc[df_membros['Nome'] == escolhido, 'escalas_no_mes'] += 1
                    else:
                        dia_escala[funcao] = "FALTA PESSOAL"
                
                escala_final.append(dia_escala)

        # Exibição dos Resultados
        df_resultado = pd.DataFrame(escala_final)
        st.subheader(f"Escala Gerada: {mes}/{ano}")
        st.dataframe(df_resultado, use_container_width=True)

        # Exportação para Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Escala')
        
        st.download_button(
            label="📥 Baixar Escala em Excel",
            data=output.getvalue(),
            file_name=f"escala_diaconato_{mes}_{ano}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Aguardando o arquivo membros_master.csv para começar.")