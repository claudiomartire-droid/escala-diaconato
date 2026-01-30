mapa_dias = {2: "Quarta_Feira", 5: "Sabado", 6: "Domingo"}
        escala_final = []
        df_membros['escalas_no_mes'] = 0 
        
        # Variável para controlar quem trabalhou no último culto (Ponto 1)
        ultimos_escalados = []
        
        for data in datas:
            data_atual = data.date()
            dia_semana_num = data.weekday()
            nome_coluna_dia = mapa_dias.get(dia_semana_num)
            
            if nome_coluna_dia in dias_culto:
                candidatos_dia = df_membros[df_membros[nome_coluna_dia] != "NÃO"].copy()
                
                # --- PONTO 1: Bloqueio de Sequência ---
                # Remove quem trabalhou no culto anterior
                candidatos_dia = candidatos_dia[~candidatos_dia['Nome'].isin(ultimos_escalados)]

                # Filtro de Ausências (Férias)
                for _, aus in df_ausencias.iterrows():
                    if pd.notna(aus['Membro']) and pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        try:
                            data_inicio = pd.to_datetime(aus['Início']).date()
                            data_fim = pd.to_datetime(aus['Fim']).date()
                            if data_inicio <= data_atual <= data_fim:
                                candidatos_dia = candidatos_dia[candidatos_dia['Nome'] != aus['Membro']]
                        except: continue

                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                escalados_no_dia = {} 

                # Definição das vagas
                if nome_coluna_dia == "Domingo":
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"]
                else:
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    # Candidatos que ainda não foram pegos para OUTRA vaga NESTE MESMO DIA
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    
                    # --- PONTO 2: Mulheres não na Portaria 1 ---
                    if vaga == "Portaria 1 (Rua)":
                        candidatos = candidatos[candidatos['Sexo'] == 'M']

                    # Regras de Duplas Impedidas
                    for _, dupla in df_duplas.iterrows():
                        if pd.notna(dupla['Pessoa A']) and pd.notna(dupla['Pessoa B']):
                            if dupla['Pessoa A'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa B']]
                            if dupla['Pessoa B'] in escalados_no_dia: candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa A']]

                    # Restrição de Função Individual (Pelo editor)
                    for _, rest in df_restricoes.iterrows():
                        if pd.notna(rest['Membro']) and pd.notna(rest['Função Proibida']) and rest['Função Proibida'] in vaga:
                            candidatos = candidatos[candidatos['Nome'] != rest['Membro']]

                    # Filtros de Gênero específicos para Frente do Templo (Domingo)
                    if "Frente Templo (M)" in vaga: candidatos = candidatos[candidatos['Sexo'] == 'M']
                    elif "Frente Templo (F)" in vaga: candidatos = candidatos[candidatos['Sexo'] == 'F']
                    
                    # Ordenação por equilíbrio de carga
                    candidatos = candidatos.sort_values(by='escalas_no_mes')

                    if not candidatos.empty:
                        escolhido = candidatos.iloc[0]
                        escalados_no_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # Atualiza a lista de "quem trabalhou por último" para o próximo culto
                ultimos_escalados = list(escalados_no_dia.keys())
                
                # Santa Ceia e Abertura (Lógica simplificada conforme versões anteriores)
                # ... [Mantém a lógica da Santa Ceia e Abertura] ...
                
                escala_final.append(dia_escala)
