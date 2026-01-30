for data in datas:
            data_atual = data.date() # Converte para data pura para comparação
            dia_semana_num = data.weekday()
            nome_coluna_dia = mapa_dias.get(dia_semana_num)
            
            if nome_coluna_dia in dias_culto:
                # 1. Filtro de Disponibilidade Geral
                candidatos_dia = df_membros[df_membros[nome_coluna_dia] != "NÃO"].copy()
                
                # --- CORREÇÃO DO ERRO DE TYPEERROR AQUI ---
                for _, aus in df_ausencias.iterrows():
                    # Garantimos que os valores de início e fim sejam convertidos para o tipo date
                    if pd.notna(aus['Início']) and pd.notna(aus['Fim']):
                        try:
                            # Converte para objeto date se for Timestamp ou String
                            inicio = aus['Início'] if isinstance(aus['Início'], date) else aus['Início'].date()
                            fim = aus['Fim'] if isinstance(aus['Fim'], date) else aus['Fim'].date()
                            
                            if inicio <= data_atual <= fim:
                                candidatos_dia = candidatos_dia[candidatos_dia['Nome'] != aus['Membro']]
                        except:
                            continue # Pula se a data estiver em formato inválido no editor
                # ------------------------------------------

                dia_escala = {"Data": data.strftime('%d/%m (%a)')}
                escalados_no_dia = {} 

                # Define as vagas (Domingo x Resto da semana)
                if nome_coluna_dia == "Domingo":
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (A)", "Portaria 2 (B)", "Frente Templo (M)", "Frente Templo (F)"]
                else:
                    vagas = ["Portaria 1 (Rua)", "Portaria 2 (Templo)", "Frente Templo"]

                for vaga in vagas:
                    candidatos = candidatos_dia[~candidatos_dia['Nome'].isin(escalados_no_dia.keys())]
                    
                    # Regra de Duplas
                    for _, dupla in df_duplas.iterrows():
                        if pd.notna(dupla['Pessoa A']) and pd.notna(dupla['Pessoa B']):
                            if dupla['Pessoa A'] in escalados_no_dia: 
                                candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa B']]
                            if dupla['Pessoa B'] in escalados_no_dia: 
                                candidatos = candidatos[candidatos['Nome'] != dupla['Pessoa A']]

                    # Regra de Restrição de Função
                    for _, rest in df_restricoes.iterrows():
                        if pd.notna(rest['Membro']) and pd.notna(rest['Função Proibida']):
                            if rest['Função Proibida'] in vaga:
                                candidatos = candidatos[candidatos['Nome'] != rest['Membro']]

                    # Filtro de Gênero para Frente do Templo no Domingo
                    if "Frente Templo (M)" in vaga: 
                        candidatos = candidatos[candidatos['Sexo'] == 'M']
                    elif "Frente Templo (F)" in vaga: 
                        candidatos = candidatos[candidatos['Sexo'] == 'F']
                    
                    # Ordenação por quem trabalhou menos no mês
                    candidatos = candidatos.sort_values(by='escalas_no_mes')

                    if not candidatos.empty:
                        escolhido = candidatos.iloc[0]
                        escalados_no_dia[escolhido['Nome']] = escolhido
                        dia_escala[vaga] = escolhido['Nome']
                        df_membros.loc[df_membros['Nome'] == escolhido['Nome'], 'escalas_no_mes'] += 1
                    else:
                        dia_escala[vaga] = "FALTA PESSOAL"

                # Santa Ceia (2H + 2M dos já escalados, exceto Portaria 1)
                if data_atual == data_ceia:
                    aptos_ceia = [m for m in escalados_no_dia.keys() if m != dia_escala.get("Portaria 1 (Rua)")]
                    # Aplica restrições de função específicas para Ceia
                    for _, rest in df_restricoes.iterrows():
                        if rest['Função Proibida'] == "Santa Ceia":
                            aptos_ceia = [m for m in aptos_ceia if m != rest['Membro']]
                    
                    homens_ceia = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'M'][:2]
                    mulheres_ceia = [m for m in aptos_ceia if escalados_no_dia[m]['Sexo'] == 'F'][:2]
                    dia_escala["Servir Santa Ceia"] = ", ".join(homens_ceia + mulheres_ceia)
                
                # Abertura (Prioriza quem já está no templo, exceto Portaria 1)
                c_abertura = candidatos_dia[(candidatos_dia['Abertura'] == "SIM") & (candidatos_dia['Nome'] != dia_escala.get("Portaria 1 (Rua)"))]
                for _, rest in df_restricoes.iterrows():
                    if rest['Função Proibida'] == "Abertura":
                        c_abertura = c_abertura[c_abertura['Nome'] != rest['Membro']]
                
                ja_no_templo = [n for n in escalados_no_dia.keys() if n in c_abertura['Nome'].values]
                if ja_no_templo:
                    dia_escala["Abertura"] = ja_no_templo[0]
                else:
                    sobras = c_abertura[~c_abertura['Nome'].isin(escalados_no_dia.keys())]
                    dia_escala["Abertura"] = sobras.iloc[0]['Nome'] if not sobras.empty else "---"
                
                escala_final.append(dia_escala)
