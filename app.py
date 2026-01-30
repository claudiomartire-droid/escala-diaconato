import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V5.6", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.6)")
st.info("⚖️ Foco: Equidade Progressiva e Multi-Histórico de Santa Ceia.")

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

# ALTERAÇÃO: Permite múltiplos arquivos de escalas passadas para consolidar o histórico
arquivos_historicos = st.sidebar.file_uploader(
    "Suba as escalas anteriores (vários arquivos)", 
    type=["csv", "xlsx"], 
    accept_multiple_files=True
)

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except Exception:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # --- PROCESSAMENTO CONSOLIDADO DE MÚLTIPLOS HISTÓRICOS ---
    contagem_ceia_historico = {nome: 0 for nome in nomes_membros}
    if arquivos_historicos:
        for arq in arquivos_historicos:
            try:
                df_h = pd.read_csv(arq) if arq.name.endswith('.csv') else pd.read_excel(arq)
                # Verifica colunas que mencionam Ceia ou Ornamentação
                cols_alvo = [c for c in df_h.columns if any(x in c for x in ["Santa Ceia", "Ornamentação"])]
                for col in cols_alvo:
                    for celula in df_h[col].dropna().astype(str):
                        for nome in nomes_membros:
                            if nome in celula:
                                contagem_ceia_historico[nome] += 1
            except:
                continue
        st.sidebar.success(f"✅ {len(arquivos_historicos)} arquivos de histórico processados!")

    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia_historico)

    # --- REGRAS (Resumo) ---
    st.subheader("📋 Painel de Controle de Equidade")
    tab_regras, tab_equidade = st.tabs(["🚫 Restrições", "🍷 Ranking de Escalação (Santa Ceia)"])
    
    with tab_equidade:
        st.write("Membros ordenados por quem serviu MENOS vezes (Prioritários para o próximo culto de Ceia):")
        # Criamos um Score: histórico antigo + escalas do mês atual
        df_view = df_membros[['Nome', 'historico_ceia']].sort_values(by='historico_ceia')
        st.dataframe(df_view, use_container_width=True)

    # --- CONFIGURAÇÕES ---
    st.sidebar.header("2. Configurações")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=ano_padrao)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=mes_padrao-1, format_func=lambda x: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes))
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    # Editor de Ausências
    ausencias = st.sidebar.data_editor(pd.DataFrame(columns=["Membro", "Início", "Fim"]), num_rows="dynamic")

    # --- MOTOR DE GERAÇÃO ---
    if st.sidebar.button("Gerar Escala com Equidade Progressiva"):
        datas_mes = pd.date_range(date(ano, mes, 1), (date(ano + (1 if mes==12 else 0), 1 if mes==12 else mes+1, 1) - timedelta(days=1)))
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        membros_ultimo_culto = []

        for data in datas_mes:
            data_atual = data.date()
            nome_col_dia = mapa_dias.get(data.weekday())
            
            if nome_col_dia in dias_semana:
                # 1. Filtro de Candidatos Disponíveis
                cands = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # 2. Aplicação de Ausências
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']):
                        if pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                            cands = cands[cands['Nome'] != aus['Membro']]

                dia_escala = {"Data": data.strftime('%d/%m/%Y (%a)')}
                escalados_dia = {}

                # 3. Definição de Vagas
                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if data.weekday() == 6 else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                # 4. Escalação das Vagas com Peso de Equidade
                for vaga in vagas:
                    vaga_cands = cands[~cands['Nome'].isin(escalados_dia.keys())]
                    if "Portaria 1" in vaga or "(M)" in vaga: vaga_cands = vaga_cands[vaga_cands['Sexo'] == 'M']
                    if "(F)" in vaga: vaga_cands = vaga_cands[vaga_cands['Sexo'] == 'F']

                    # SE FOR CEIA: O critério principal é o HISTÓRICO ACUMULADO
                    if data_atual == data_ceia:
                        vaga_cands = vaga_cands.sort_values(by=['historico_ceia', 'escalas_no_mes'])
                    else:
                        vaga_cands = vaga_cands.sort_values(by='escalas_no_mes')

                    if not vaga_cands.empty:
                        escolhido = vaga_cands.iloc[0]
                        dia_escala[vaga] = escolhido['Nome']
                        escalados_dia[escolhido['Nome']] = escolhido
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # 5. Santa Ceia (Servir e Ornamentação)
                if data_atual == data_ceia:
                    # Ornamentação (Quem serviu menos na história)
                    aptos_orn = cands[(cands['Ornamentacao'] == "SIM") & (~cands['Nome'].isin(escalados_dia.keys()))]
                    escolhidos_orn = aptos_orn.sort_values(by=['historico_ceia', 'escalas_no_mes']).head(2)
                    dia_escala["Ornamentação"] = ", ".join(escolhidos_orn['Nome'].tolist())
                    for n in escolhidos_orn['Nome']:
                        df_membros.loc[df_membros['Nome'] == n, 'escalas_no_mes'] += 0.5
                    
                    # Servir Ceia (Entre os já escalados no dia, priorizar rodízio histórico)
                    def get_hist(n): return df_membros.loc[df_membros['Nome'] == n, 'historico_ceia'].values[0]
                    escalados_nomes = sorted(list(escalados_dia.keys()), key=get_hist)
                    dia_escala["Servir Santa Ceia"] = ", ".join(escalados_nomes[:4])

                escala_final.append(dia_escala)
                membros_ultimo_culto = list(escalados_dia.keys())

        # Exibição e Download
        df_res = pd.DataFrame(escala_final)
        st.subheader("🗓️ Escala Mensal Gerada")
        st.dataframe(df_res, use_container_width=True)

        # --- NOVA FUNCIONALIDADE: EXPORTAR COM HISTÓRICO ATUALIZADO ---
        # Criamos um DataFrame de "Memória" para ser usado no próximo mês
        df_memoria = df_membros[['Nome', 'historico_ceia']].copy()
        # Se a pessoa serviu na ceia hoje, somamos +1 no histórico dela para o arquivo de saída
        for nome in nomes_membros:
            foi_na_ceia = False
            # Verifica se o nome aparece na linha da Santa Ceia da escala gerada
            linha_ceia = df_res[df_res['Data'].str.contains(data_ceia.strftime('%d/%m/%Y'))]
            if not linha_ceia.empty:
                txt_ceia = str(linha_ceia.iloc[0].to_dict())
                if nome in txt_ceia:
                    df_memoria.loc[df_memoria['Nome'] == nome, 'historico_ceia'] += 1

        col1, col2 = st.columns(2)
        with col1:
            out_escala = io.BytesIO()
            df_res.to_excel(out_escala, index=False)
            st.download_button("📥 Baixar Escala (Excel)", out_escala.getvalue(), f"escala_{mes}_{ano}.xlsx")
        
        with col2:
            out_hist = io.BytesIO()
            df_memoria.to_csv(out_hist, index=False)
            st.download_button("💾 Baixar NOVO Histórico Atualizado", out_hist.getvalue(), "historico_consolidado.csv", help="Suba este arquivo no próximo mês para manter o rodízio perfeito!")

else:
    st.info("Aguardando arquivo membros_master.csv para iniciar.")
