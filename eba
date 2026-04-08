def corrigir_duplicidade_sped(arquivo_entrada, arquivo_saida):
    with open(arquivo_entrada, 'r', encoding='latin-1') as f:
        linhas = f.readlines()

    novas_linhas = []
    i = 0
    n = len(linhas)
    
    # Dicionário para rastrear notas processadas (Chave da NF ou Número como identificador)
    # No seu exemplo, o campo da chave é o 9º campo (índice 9)
    notas_mantidas = set()

    while i < n:
        linha_atual = linhas[i]
        campos = linha_atual.split('|')

        if len(campos) > 1 and campos[1] == 'C100':
            chave_nf = campos[9] # Campo da Chave de Acesso
            
            # Verifica se a próxima linha é um C170
            tem_c170_seguida = False
            if i + 1 < n:
                proxima_linha = linhas[i+1].split('|')
                if len(proxima_linha) > 1 and proxima_linha[1] == 'C170':
                    tem_c170_seguida = True

            # Lógica de exclusão:
            # Se a nota já foi processada ou não tem C170, e existe uma duplicata com C170, filtramos.
            if tem_c170_seguida:
                novas_linhas.append(linha_atual)
                notas_mantidas.add(chave_nf)
            else:
                # Se não tem C170, só adicionamos se a chave ainda não existe no arquivo
                # (caso seja uma nota de serviço ou sem itens propositalmente)
                if chave_nf not in notas_mantidas:
                    # Precisamos checar se não há uma versão com C170 mais adiante no arquivo
                    # Para simplificar: se não tem C170 agora, mas é duplicada, nós pulamos.
                    duplicada = False
                    for j in range(i + 1, min(i + 50, n)): # Busca curta por duplicata
                        outra_linha = linhas[j].split('|')
                        if len(outra_linha) > 9 and outra_linha[1] == 'C100' and outra_linha[9] == chave_nf:
                            # Se achou outra C100 igual, vamos ver se a outra tem C170
                            if j + 1 < n and linhas[j+1].split('|')[1] == 'C170':
                                duplicada = True
                                break
                    
                    if not duplicada:
                        novas_linhas.append(linha_atual)
                        notas_mantidas.add(chave_nf)
        else:
            # Mantém todos os outros registros (0000, C110, C190, Reg M, etc)
            novas_linhas.append(linha_atual)
        
        i += 1

    with open(arquivo_saida, 'w', encoding='latin-1') as f:
        f.writelines(novas_linhas)
    
    print(f"Arquivo corrigido salvo em: {arquivo_saida}")

# Modo de usar:
# corrigir_duplicidade_sped('seu_arquivo_original.txt', 'arquivo_limpo.txt')
