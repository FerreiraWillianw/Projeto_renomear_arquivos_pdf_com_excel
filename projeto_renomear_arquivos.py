import os
import pandas as pd
import time
from datetime import datetime

# --- Configurações Principais (Ajuste Apenas Aqui!) ---

# 1. Caminho Fixo para a Planilha Principal
# Use 'r' antes da string para garantir que as barras invertidas sejam tratadas corretamente (ex: r"C:\Meu Caminho\Arquivo.xlsx")
CAMINHO_PARA_MINHA_PLANILHA = r"y:\2.Planilhas Principais\PLANILHA GERAL 3.xlsx"

# 2. Nome Fixo da Aba (Sheet) dentro da Planilha Principal
NOME_DA_ABA_PLANILHA = "Processo" # Nome EXATO da aba (case-sensitive!)

# 3. Nomes EXATOS das Colunas na Sua Planilha (Case-Sensitive!)
COLUNA_NUMERO_PROCESSO = "Número do processo"
COLUNA_NOME_CLIENTE = "Cliente principal"
COLUNA_PARTE_ADVERSA = "Contrário principal"
COLUNA_NOME_ADVOGADO = "ADVOGADO"

# 4. Parâmetros para tentativas de renomeação em caso de erro (arquivo em uso, etc.)
NUM_TENTATIVAS = 3      # Tentar renomear 3 vezes
ATRASO_SEGUNDOS = 1     # Esperar 1 segundo entre as tentativas

# --- Fim das Configurações ---


def renomear_arquivos_por_planilha(
    caminho_planilha,
    nome_aba_planilha,
    coluna_processo_planilha,
    coluna_cliente_planilha,
    coluna_parte_adversa_planilha,
    coluna_advogado_planilha,
    pasta_arquivos,
    max_tentativas=NUM_TENTATIVAS,
    atraso_entre_tentativas=ATRASO_SEGUNDOS
):
    """
    Renomeia arquivos PDF em uma pasta específica com base em dados de uma planilha Excel.

    Args:
        caminho_planilha (str): Caminho completo para o arquivo Excel.
        nome_aba_planilha (str): Nome da aba dentro da planilha onde os dados estão.
        coluna_processo_planilha (str): Nome da coluna com o número do processo.
        coluna_cliente_planilha (str): Nome da coluna com o nome do cliente.
        coluna_parte_adversa_planilha (str): Nome da coluna com a parte adversa.
        coluna_advogado_planilha (str): Nome da coluna com o nome do advogado.
        pasta_arquivos (str): Caminho completo da pasta que contém os arquivos a serem renomeados.
        max_tentativas (int): Número máximo de tentativas para renomear um arquivo.
        atraso_entre_tentativas (int): Atraso em segundos entre as tentativas de renomeação.
    """
    print(f"\n--- Iniciando o processo de renomeação de arquivos ---")
    print(f"Planilha de referência: '{caminho_planilha}'")
    print(f"Aba da planilha: '{nome_aba_planilha}'")
    print(f"Pasta de arquivos a serem renomeados: '{pasta_arquivos}'")
    print("-" * 50)

    try:
        # 1. Preparar a lista de colunas necessárias para otimizar a leitura do Excel
        colunas_necessarias = [
            coluna_processo_planilha,
            coluna_cliente_planilha,
            coluna_parte_adversa_planilha,
            coluna_advogado_planilha
        ]

        # --- NOVA VALIDAÇÃO: Verifica se a planilha existe antes de tentar ler ---
        if not os.path.exists(caminho_planilha):
            raise FileNotFoundError(f"A planilha '{caminho_planilha}' não foi encontrada. Por favor, verifique o caminho.")
        # --- FIM DA NOVA VALIDAÇÃO ---

        # 2. Carregar a planilha Excel com Pandas
        print(f"Lendo a planilha '{caminho_planilha}' na aba '{nome_aba_planilha}'...")
        df = pd.read_excel(
            caminho_planilha,
            sheet_name=nome_aba_planilha,
            usecols=colunas_necessarias
        )
        print(f"Planilha carregada com sucesso. Total de {len(df)} linhas lidas.")

        # Remover linhas onde o número do processo é nulo/vazio
        df.dropna(subset=[coluna_processo_planilha], inplace=True)
        if df.empty:
            print("\nAVISO: Nenhuma linha válida encontrada na coluna 'Número do processo' após remover vazias.")
            return

        # Verificar e tratar duplicatas no número do processo
        if df[coluna_processo_planilha].duplicated().any():
            print("\nAVISO: Duplicatas encontradas na coluna 'Número do processo'. Removendo para evitar conflitos.")
            # Manter apenas a primeira ocorrência de cada número de processo único
            df_deduplicado = df.drop_duplicates(subset=[coluna_processo_planilha], keep='first')
            print(f"Após remoção de duplicatas, restaram {len(df_deduplicado)} linhas únicas.")
        else:
            df_deduplicado = df.copy()

        # Criar um dicionário para busca rápida dos dados do processo
        # O número do processo será a chave, e o valor será um dicionário com os dados da linha
        dados_processos = df_deduplicado.set_index(coluna_processo_planilha).to_dict('index')
        print(f"Dicionário de processos criado. {len(dados_processos)} processos únicos para busca.")

        # 3. Listar arquivos na pasta
        if not os.path.exists(pasta_arquivos):
            raise FileNotFoundError(f"A pasta '{pasta_arquivos}' não existe. Por favor, verifique o caminho informado.")

        if not os.path.isdir(pasta_arquivos):
            raise NotADirectoryError(f"O caminho '{pasta_arquivos}' não é uma pasta válida. Por favor, verifique.")

        arquivos_na_pasta = [f for f in os.listdir(pasta_arquivos) if os.path.isfile(os.path.join(pasta_arquivos, f))]
        print(f"Encontrados {len(arquivos_na_pasta)} arquivos na pasta '{pasta_arquivos}'.")

        if not arquivos_na_pasta:
            print("\nAVISO: NENHUM arquivo encontrado na pasta especificada. Nada para renomear.")
            return

        # 4. Iterar sobre os arquivos e renomear
        arquivos_renomeados_com_sucesso = 0
        arquivos_nao_encontrados_na_planilha = 0
        arquivos_com_erro_renomeacao = 0

        print("\nIniciando renomeação dos arquivos...")
        for nome_arquivo_antigo in arquivos_na_pasta:
            caminho_arquivo_antigo = os.path.join(pasta_arquivos, nome_arquivo_antigo)
            nome_base, extensao = os.path.splitext(nome_arquivo_antigo)

            # Assumimos que o nome do arquivo (sem extensão) é o número do processo
            numero_processo_arquivo = nome_base.strip()

            if numero_processo_arquivo in dados_processos:
                dados = dados_processos[numero_processo_arquivo]

                # Usar .get() com um valor padrão para evitar KeyError se a coluna estiver vazia na planilha
                nome_advogado = str(dados.get(coluna_advogado_planilha, "ADVOGADO_NAO_INFO")).strip()
                nome_cliente = str(dados.get(coluna_cliente_planilha, "CLIENTE_NAO_INFO")).strip()
                parte_adversa = str(dados.get(coluna_parte_adversa_planilha, "ADVERSO_NAO_INFO")).strip()

                # Construir o novo nome do arquivo
                novo_nome_arquivo = f"{numero_processo_arquivo} - {nome_advogado} - {nome_cliente} X {parte_adversa}{extensao}"

                # Limpar o novo nome para remover caracteres inválidos em nomes de arquivo
                # Caracteres que geralmente não são permitidos em nomes de arquivo do Windows: \ / : * ? " < > |
                caracteres_invalidos = r'\/|:*?"<>'
                for char in caracteres_invalidos:
                    novo_nome_arquivo = novo_nome_arquivo.replace(char, '')

                # Remover múltiplos espaços e espaços no início/fim
                novo_nome_arquivo = ' '.join(novo_nome_arquivo.split()).strip()

                caminho_arquivo_novo = os.path.join(pasta_arquivos, novo_nome_arquivo)

                if caminho_arquivo_antigo == caminho_arquivo_novo:
                    print(f"IGNORANDO: Arquivo '{nome_arquivo_antigo}' já está com o nome correto.")
                    arquivos_renomeados_com_sucesso += 1 # Contar como sucesso pois já está ok
                    continue

                renomeado = False
                for tentativa in range(1, max_tentativas + 1):
                    try:
                        os.rename(caminho_arquivo_antigo, caminho_arquivo_novo)
                        print(f"Renomeado: '{nome_arquivo_antigo}' -> '{novo_nome_arquivo}'")
                        arquivos_renomeados_com_sucesso += 1
                        renomeado = True
                        break   # Sai do loop de tentativas se for bem-sucedido
                    except OSError as e:
                        if tentativa < max_tentativas:
                            print(f"AVISO: Não foi possível renomear '{nome_arquivo_antigo}'. Erro: {e}. Tentando novamente em {atraso_entre_tentativas}s...")
                            time.sleep(atraso_entre_tentativas)
                        else:
                            print(f"ERRO: Falha ao renomear '{nome_arquivo_antigo}' após {max_tentativas} tentativas. Erro final: {e}")
                            arquivos_com_erro_renomeacao += 1
                if not renomeado and not (caminho_arquivo_antigo == caminho_arquivo_novo):
                    arquivos_com_erro_renomeacao += 1 # Garante que o erro seja contado se não renomeou e não era o mesmo nome
            else:
                print(f"AVISO: Arquivo '{nome_arquivo_antigo}' não encontrado na planilha. Não será renomeado.")
                arquivos_nao_encontrados_na_planilha += 1

        print("\n--- Processo de Renomeação Concluído ---")
        print(f"Total de arquivos na pasta: {len(arquivos_na_pasta)}")
        print(f"Arquivos renomeados/já corretos: {arquivos_renomeados_com_sucesso}")
        print(f"Arquivos não encontrados na planilha: {arquivos_nao_encontrados_na_planilha}")
        print(f"Arquivos com erro de renomeação: {arquivos_com_erro_renomeacao}")

    except FileNotFoundError as e:
        print(f"\nERRO FATAL: Arquivo ou pasta não encontrado(a). Detalhes: {e}")
        # if "planilha" in str(e).lower() or "excel" in str(e).lower(): # Esta linha foi generalizada pela nova validação acima
        #     print(f"Verifique se o caminho da planilha '{caminho_planilha}' está correto.")
        # elif "pasta" in str(e).lower() or "directory" in str(e).lower(): # Esta linha foi generalizada pela nova validação acima
        #     print(f"Verifique se o caminho da pasta de arquivos '{pasta_arquivos}' está correto.")
    except NotADirectoryError as e: # Captura o novo erro específico para caminho não ser diretório
        print(f"\nERRO FATAL: O caminho informado para a pasta de arquivos não é um diretório válido. Detalhes: {e}")
        print(f"Verifique se o caminho '{pasta_arquivos}' realmente aponta para uma pasta.")
    except KeyError as e:
        print(f"\nERRO FATAL: Coluna não encontrada na planilha. Detalhes: {e}")
        print(f"Por favor, verifique se os nomes das colunas nas configurações (COLUNA_NUMERO_PROCESSO, COLUNA_NOME_CLIENTE, etc.) são EXATOS e case-sensitive na planilha.")
    except ValueError as e:
        print(f"\nERRO FATAL: Erro ao carregar a planilha. Detalhes: {e}")
        if "sheet" in str(e).lower():
            print(f"Verifique se o nome da aba '{nome_aba_planilha}' está correto na planilha e se ela existe.")
    except pd.errors.EmptyDataError:
        print(f"\nERRO FATAL: A planilha '{caminho_planilha}' está vazia ou não contém dados na aba '{nome_aba_planilha}'.")
    except Exception as e:
        print(f"\nERRO FATAL INESPERADO: Ocorreu um erro não tratado. Detalhes: {e}")

# --- Execução Principal do Script ---
if __name__ == "__main__":
    # --- Parte de Configuração de Teste (OPCIONAL: apenas para criar pastas e arquivos de teste) ---
    # Para usar, descomente as linhas abaixo e rode o script APENAS UMA VEZ para criar os arquivos.
    # Lembre-se de re-comentar ANTES de usar o script para renomear seus arquivos reais!

    # print("\n--- ATENÇÃO: CRIANDO ESTRUTURA DE TESTE ---")
    # try:
    #     PASTA_TESTE_ROOT = "Y:\\TEMP_RENOMEADOR_TESTE"
    #     os.makedirs(os.path.join(PASTA_TESTE_ROOT, "2025", "5. maio", "29 05 2025"), exist_ok=True)
    #     pasta_arquivos_teste = os.path.join(PASTA_TESTE_ROOT, "2025", "5. maio", "29 05 2025")
    #     planilha_teste_path = os.path.join(PASTA_TESTE_ROOT, "PLANILHA_TESTE.xlsx")
    #
    #     # Cria arquivos de teste
    #     open(os.path.join(pasta_arquivos_teste, "12345.pdf"), "w").close()
    #     open(os.path.join(pasta_arquivos_teste, "67890.pdf"), "w").close()
    #     open(os.path.join(pasta_arquivos_teste, "99999.pdf"), "w").close() # Arquivo não na planilha
    #     open(os.path.join(pasta_arquivos_teste, "11111.pdf"), "w").close() # Duplicata para simular
    #
    #     # Cria planilha de teste
    #     dados_teste = {
    #         "Número do processo": ["12345", "67890", "11111", "22222"],
    #         "Cliente principal": ["Empresa A", "João Silva", "Empresa A", "Maria Souza"],
    #         "Contrário principal": ["Empresa B", "Maria Oliveira", "Empresa B", "Carlos Pereira"],
    #         "ADVOGADO": ["Dr. Fulano", "Dra. Ciclana", "Dr. Fulano", "Dra. Beltrana"]
    #     }
    #     df_teste = pd.DataFrame(dados_teste)
    #     df_teste.to_excel(planilha_teste_path, sheet_name="Processo", index=False)
    #
    #     print(f"Estrutura de teste criada em: {PASTA_TESTE_ROOT}")
    #     print("ATENÇÃO: Descomente as linhas abaixo e use os caminhos de teste para rodar o script.")
    #     print("-----------------------------------------------------------------------------------\n")
    #
    #     # Se quiser usar os caminhos de teste, descomente estas linhas
    #     # CAMINHO_PARA_MINHA_PLANILHA = planilha_teste_path
    #     # PASTA_DOS_MEUS_ARQUIVOS_REAL = pasta_arquivos_teste # Renomeada para evitar conflito com input()
    # except Exception as e:
    #     print(f"ERRO ao criar estrutura de teste: {e}")
    #     print("Verifique se você tem permissão para criar pastas no drive Y: ou ajuste o CAMINHO_TESTE_ROOT.")
    #
    # print("-" * 50)
    # --- Fim da Parte de Configuração de Teste ---

    # --- Pede o caminho da pasta de arquivos ao usuário ---
    # Coloque a validação para o usuário digitar um caminho válido
    PASTA_DOS_MEUS_ARQUIVOS_INFORMADA = input("Por favor, digite o caminho COMPLETO da pasta que contém os arquivos a serem renomeados (ex: C:\\Pasta\\Subpasta): ")

    # Loop de validação do caminho
    while not os.path.isdir(PASTA_DOS_MEUS_ARQUIVOS_INFORMADA):
        print(f"Erro: O caminho '{PASTA_DOS_MEUS_ARQUIVOS_INFORMADA}' não foi encontrado ou não é uma pasta válida.")
        PASTA_DOS_MEUS_ARQUIVOS_INFORMADA = input("Por favor, digite um caminho VÁLIDO para a pasta dos arquivos: ")

    # Chama a função principal com as configurações definidas
    renomear_arquivos_por_planilha(
        caminho_planilha=CAMINHO_PARA_MINHA_PLANILHA,
        nome_aba_planilha=NOME_DA_ABA_PLANILHA,
        coluna_processo_planilha=COLUNA_NUMERO_PROCESSO,
        coluna_cliente_planilha=COLUNA_NOME_CLIENTE,
        coluna_parte_adversa_planilha=COLUNA_PARTE_ADVERSA,
        coluna_advogado_planilha=COLUNA_NOME_ADVOGADO,
        pasta_arquivos=PASTA_DOS_MEUS_ARQUIVOS_INFORMADA # Usa o caminho informado pelo usuário
    )

    # --- NOVO: Mantém o Prompt Aberto ---
    print("\n--- Pressione ENTER para fechar a janela ---")
    input()
    # --- FIM DO NOVO ---
