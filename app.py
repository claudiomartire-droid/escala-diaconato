import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm # Para gerenciar fontes
import numpy as np # Para calcular o comprimento da coluna para ajuste

# Configuração da Página
st.set_page_config(page_title="Gerador de Escala Diaconato V5.6", layout="wide")

st.title("⛪ Gerador de Escala de Diaconato (Versão 5.6)")

# --- INICIALIZAÇÃO DO ESTADO ---
if 'escala_gerada' not in st.session_state:
    st.session_state.escala_gerada = None
if 'df_memoria' not in st.session_state:
    st.session_state.df_memoria = None

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

# --- 1. CARGA DE DADOS (SIDEBAR) ---
st.sidebar.header("1. Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba o arquivo membros_master.csv", type="csv")
arquivos_historicos = st.sidebar.file_uploader("Suba históricos ou consolidado", type=["csv", "xlsx"], accept_multiple_files=True)

if arquivo_carregado:
    try:
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='iso-8859-1')
    except Exception:
        arquivo_carregado.seek(0)
        df_membros = pd.read_csv(arquivo_carregado, sep=None, engine='python', encoding='utf-8-sig')

    nomes_membros = sorted(df_membros['Nome'].tolist())
    
    # Processamento de Equidade Contínua
    contagem_ceia_historico = {nome: 0 for nome in nomes_membros}
    if arquivos_historicos:
        for arq in arquivos_historicos:
            try:
                df_h = pd.read_csv(arq) if arq.name.endswith('.csv') else pd.read_excel(arq)
                if 'historico_ceia' in df_h.columns and 'Nome' in df_h.columns:
                    for _, row in df_h.iterrows():
                        if row['Nome'] in contagem_ceia_historico:
                            contagem_ceia_historico[row['Nome']] += row['historico_ceia']
                else:
                    cols_alvo = [c for c in df_h.columns if any(x in c for x in ["Santa Ceia", "Ornamentação"])]
                    for col in cols_alvo:
                        for celula in df_h[col].dropna().astype(str):
                            for nome in nomes_membros:
                                if nome in celula: contagem_ceia_historico[nome] += 1
            except: pass # Ignora erros de leitura de arquivos incompletos

    df_membros['historico_ceia'] = df_membros['Nome'].map(contagem_ceia_historico)

    # Regras de Duplas e Funções
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

    # TABS DE CONFERÊNCIA
    st.subheader("📋 Conferência de Regras e Equidade")
    t1, t2, t3 = st.tabs(["👥 Duplas Impedidas", "🚫 Restrições de Função", "🍷 Ranking Santa Ceia"])
    with t1: 
        if regras_duplas: st.dataframe(pd.DataFrame(regras_duplas), use_container_width=True)
        else: st.info("Sem duplas impeditivas.")
    with t2: 
        if regras_funcao: st.dataframe(pd.DataFrame(regras_funcao), use_container_width=True)
        else: st.info("Sem restrições de função.")
    with t3: 
        st.dataframe(df_membros[['Nome', 'historico_ceia']].sort_values(by='historico_ceia'), use_container_width=True)

    # --- 2. CONFIGURAÇÕES (SIDEBAR) ---
    st.sidebar.header("2. Configurações")
    ano = st.sidebar.number_input("Ano", min_value=2025, max_value=2030, value=ano_padrao)
    mes = st.sidebar.selectbox("Mês", range(1, 13), index=mes_padrao-1, format_func=lambda x: ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][x-1])
    dias_semana = st.sidebar.multiselect("Dias de Culto", ["Quarta_Feira", "Sabado", "Domingo"], default=["Quarta_Feira", "Sabado", "Domingo"])
    
    data_inicio_mes = date(ano, mes, 1)
    data_fim_mes = (date(ano + (1 if mes==12 else 0), 1 if mes==12 else mes+1, 1) - timedelta(days=1))
    datas_excluir = st.sidebar.multiselect("Datas para EXCLUIR", options=pd.date_range(data_inicio_mes, data_fim_mes), format_func=lambda x: x.strftime('%d/%m/%Y'))
    data_ceia = st.sidebar.date_input("Data da Santa Ceia", value=obter_primeiro_domingo(ano, mes))

    # --- 3. FÉRIAS / AUSÊNCIAS (CALENDÁRIO CORRIGIDO) ---
    st.sidebar.header("3. Férias / Ausências")
    ausencias = st.sidebar.data_editor(
        pd.DataFrame(columns=["Membro", "Início", "Fim"]),
        column_config={
            "Membro": st.column_config.SelectboxColumn(options=nomes_membros, required=True),
            "Início": st.column_config.DateColumn(format="DD/MM/YYYY", required=True),
            "Fim": st.column_config.DateColumn(format="DD/MM/YYYY", required=True)
        },
        num_rows="dynamic"
    )

    # --- 4. MOTOR DE GERAÇÃO ---
    if st.sidebar.button("Gerar Escala Atualizada"):
        datas_mes = pd.date_range(data_inicio_mes, data_fim_mes)
        mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0.0
        membros_ultimo_culto = []

        for data in datas_mes:
            data_atual = data.date()
            if any(data_atual == d.date() for d in datas_excluir): continue
            
            nome_col_dia = mapa_dias.get(data.weekday())
            if nome_col_dia in dias_semana:
                cands = df_membros[df_membros[nome_col_dia] != "NÃO"].copy()
                cands = cands[~cands['Nome'].isin(membros_ultimo_culto)]
                
                # Filtro de Ausências
                for _, aus in ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        if pd.to_datetime(aus['Início']).date() <= data_atual <= pd.to_datetime(aus['Fim']).date():
                            cands = cands[cands['Nome'] != aus['Membro']]

                dia_escala = {"Data": data.strftime('%d/%m/%Y (%a)')}
                escalados_dia = {}
                vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"] if data.weekday() == 6 else ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                # Escala de Portarias/Frente
                for vaga in vagas:
                    v_cands = cands[~cands['Nome'].isin(escalados_dia.keys())]
                    if "Portaria 1" in vaga or "(M)" in vaga: v_cands = v_cands[v_cands['Sexo'] == 'M']
                    if "(F)" in vaga: v_cands = v_cands[v_cands['Sexo'] == 'F']
                    
                    for r in regras_duplas:
                        if r['Membro'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Evitar Escalar Com']]
                        if r['Evitar Escalar Com'] in escalados_dia: v_cands = v_cands[v_cands['Nome'] != r['Membro']]
                    for rest in regras_funcao:
                        if rest['Função Proibida'] in vaga: v_cands = v_cands[v_cands['Nome'] != rest['Membro']]

                    # Ordenação de Equidade
                    ordem = ['historico_ceia', 'escalas_no_mes'] if data_atual == data_ceia else ['escalas_no_mes']
                    v_cands = v_cands.sort_values(by=ordem)

                    if not v_cands.empty:
                        escolhido = v_cands.iloc[0]
                        dia_escala[vaga] = escolhido['Nome']
                        escalados_dia[escolhido['Nome']] = escolhido
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else: dia_escala[vaga] = "FALTA PESSOAL"

                # RESTAURADO: FUNÇÃO ABERTURA (Peso 0.5)
                aptos_ab = cands[cands['Abertura'] == "SIM"].copy()
                aptos_ab = aptos_ab[~aptos_ab['Nome'].isin([r['Membro'] for r in regras_funcao if r['Função Proibida'] == "Abertura"])]
                
                ja_no_dia = [n for n in escalados_dia.keys() if n in aptos_ab['Nome'].values and n != dia_escala.get("Portaria 1 (Rua)")]
                if ja_no_dia:
                    dia_escala["Abertura"] = ja_no_dia[0]
                else:
                    sobra_ab = aptos_ab[~aptos_ab['Nome'].isin(escalados_dia.keys())]
                    if not sobra_ab.empty:
                        escolhido_ab = sobra_ab.sort_values(by=['historico_ceia', 'escalas_no_mes']).iloc[0]
                        dia_escala["Abertura"] = escolhido_ab['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido_ab['Nome'], 'escalas_no_mes'] += 0.5
                        escalados_dia[escolhido_ab['Nome']] = escolhido_ab
                    else: dia_escala["Abertura"] = "---"

                # Santa Ceia e Ornamentação
                if data_atual == data_ceia:
                    aptos_orn = cands[(cands['Ornamentacao'] == "SIM") & (~cands['Nome'].isin(escalados_dia.keys()))]
                    esc_orn = aptos_orn.sort_values(by=['historico_ceia', 'escalas_no_mes']).head(2)
                    dia_escala["Ornamentação"] = ", ".join(esc_orn['Nome'].tolist())
                    for n in esc_orn['Nome']: df_membros.loc[df_membros['Nome'] == n, 'escalas_no_mes'] += 0.5
                    
                    esc_nomes = sorted(list(escalados_dia.keys()), key=lambda x: df_membros.loc[df_membros['Nome'] == x, 'historico_ceia'].values[0])
                    dia_escala["Servir Santa Ceia"] = ", ".join(esc_nomes[:4])

                escala_final.append(dia_escala)
                membros_ultimo_culto = list(escalados_dia.keys())
        
        st.session_state.escala_gerada = pd.DataFrame(escala_final)
        
        # Consolidação do Memória
        df_mem = df_membros[['Nome', 'historico_ceia']].copy()
        for nome in nomes_membros:
            linha_ceia = st.session_state.escala_gerada[st.session_state.escala_gerada['Data'].str.contains(data_ceia.strftime('%d/%m/%Y'))]
            if not linha_ceia.empty and nome in str(linha_ceia.iloc[0].to_dict()):
                df_mem.loc[df_mem['Nome'] == nome, 'historico_ceia'] += 1
        st.session_state.df_memoria = df_mem

    # --- ÁREA DE EXIBIÇÃO E EXPORTAÇÃO FORMATADA ---
    if st.session_state.escala_gerada is not None:
        st.subheader("🗓️ Escala Mensal Gerada")
        st.dataframe(st.session_state.escala_gerada, use_container_width=True)
        
        # Cria um contêiner para os botões de download
        col_excel, col_image, col_history = st.columns(3)
        
        with col_excel:
            # --- LÓGICA DE EXPORTAÇÃO EXCEL COM FORMATAÇÃO ---
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                df_res = st.session_state.escala_gerada
                df_res.to_excel(writer, index=False, sheet_name='Escala')
                
                workbook  = writer.book
                worksheet = writer.sheets['Escala']

                # Definição de Formatos
                header_format = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'valign': 'vcenter',
                    'align': 'center', 'fg_color': '#1F4E78', 'font_color': 'white', 'border': 1
                })
                cell_format = workbook.add_format({
                    'valign': 'vcenter', 'align': 'center', 'border': 1
                })
                ceia_highlight = workbook.add_format({
                    'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'
                })

                # Aplicar cabeçalho formatado
                for col_num, value in enumerate(df_res.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                # Aplicar formatos nas linhas e destacar Santa Ceia
                for row_num in range(len(df_res)):
                    # Verifica se a data atual da linha é a data da Santa Ceia
                    data_str_na_celula = str(df_res.iloc[row_num, 0]).split(' ')[0] # Pega 'DD/MM/YYYY'
                    is_ceia = data_ceia.strftime('%d/%m/%Y') == data_str_na_celula
                    current_format = ceia_highlight if is_ceia else cell_format
                    
                    for col_num in range(len(df_res.columns)):
                        worksheet.write(row_num + 1, col_num, df_res.iloc[row_num, col_num], current_format)

                # Ajuste Automático da Largura das Colunas
                for i, col in enumerate(df_res.columns):
                    # Calcula o comprimento máximo da string na coluna, incluindo o cabeçalho
                    max_len = max(df_res[col].astype(str).str.len().max(), len(col))
                    # Define a largura da coluna (adiciona um pouco de padding)
                    worksheet.set_column(i, i, max_len + 2)


            st.download_button(
                label="📥 Baixar Escala Formatada (Excel)",
                data=output_excel.getvalue(),
                file_name=f"Escala_Diaconato_{mes}_{ano}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="d_esc_pretty"
            )

        with col_image:
            # --- LÓGICA DE GERAÇÃO DE IMAGEM DA ESCALA ---
            df_to_image = st.session_state.escala_gerada.copy()
            
            # Formatação para o dia da semana na imagem
            df_to_image['Data'] = df_to_image['Data'].apply(lambda x: x.replace(' (Qua)', ' (Quarta)').replace(' (Sab)', ' (Sábado)').replace(' (Dom)', ' (Domingo)'))

            # Configurações de estilo para a imagem
            # Tentativa de usar fonte mais amigável, mas Streamlit/Matplotlib podem ter restrições de fonte em ambientes cloud
            # font_path = 'path/to/your/font.ttf' # Opcional: Especifique um caminho se tiver uma fonte customizada
            # if os.path.exists(font_path):
            #     fm.fontManager.addfont(font_path)
            #     prop = fm.FontProperties(fname=font_path)
            #     plt.rcParams['font.family'] = prop.get_name()
            # else:
            plt.rcParams['font.family'] = 'DejaVu Sans' # Fonte padrão do Matplotlib

            fig, ax = plt.subplots(figsize=(len(df_to_image.columns) * 2.5, len(df_to_image) * 0.7)) # Ajusta tamanho da figura
            ax.axis('off') # Remove os eixos do gráfico

            # Estilos de tabela
            table = ax.table(
                cellText=df_to_image.values,
                colLabels=df_to_image.columns,
                cellLoc='center',
                loc='center',
                colColours=['#1F4E78'] * len(df_to_image.columns) # Cor do cabeçalho
            )
            
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2) # Aumenta um pouco o tamanho das células

            # Formatação das células
            for (i, j), cell in table.get_celld().items():
                cell.set_edgecolor('#B0B0B0') # Bordas cinzas
                cell.set_linewidth(1)
                
                if i == 0: # Cabeçalho
                    cell.set_text_props(weight='bold', color='white')
                else: # Corpo da tabela
                    cell.set_text_props(color='black')
                    
                    # Destacar a linha da Santa Ceia
                    data_celula_str = df_to_image.iloc[i-1, 0].split(' ')[0] # Pega 'DD/MM/YYYY' da coluna Data
                    if data_ceia.strftime('%d/%m/%Y') == data_celula_str:
                        cell.set_facecolor('#D9E1F2') # Fundo cinza claro para Santa Ceia

            # Ajuste de layout para evitar cortes
            fig.tight_layout()
            
            # Salva a imagem em um buffer
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', pad_inches=0.5, dpi=300)
            buf.seek(0)
            plt.close(fig) # Fecha a figura para liberar memória

            st.download_button(
                label="📸 Baixar Imagem da Escala (PNG)",
                data=buf.getvalue(),
                file_name=f"Escala_Diaconato_{mes}_{ano}.png",
                mime="image/png",
                key="d_img_v5"
            )

        with col_history:
            # Exportação Simples do Histórico
            out_h = io.BytesIO()
            st.session_state.df_memoria.to_csv(out_h, index=False)
            st.download_button(
                label="💾 Baixar Histórico Acumulado",
                data=out_h.getvalue(),
                file_name="historico_consolidado.csv",
                key="d_hist_v5"
            )
else:
    st.info("Aguardando arquivo membros_master.csv.")
