import os
import pandas as pd
from tqdm import tqdm
import volumetria 

PASTA_PROJETO = r'C:\Users\nikolas.alexandre\OneDrive - unicesumar.edu.br\Área de Trabalho\Bot Integrado'
DIRETORIO_ENTRADA = os.path.join(PASTA_PROJETO, 'Relatorios Originais')
DIRETORIO_SAIDA = os.path.join(PASTA_PROJETO, 'Relatorios Convertidos')

def converter_arquivo(caminho_entrada, diretorio_saida):
    nome_base_arquivo = os.path.basename(caminho_entrada)
    print(f"\nProcessando: '{nome_base_arquivo}'")

    codificacoes = ['utf-8', 'latin-1', 'cp1252']
    html_content = None
    for encoding in codificacoes:
        try:
            with open(caminho_entrada, 'r', encoding=encoding) as f: html_content = f.read()
            break
        except UnicodeDecodeError: continue
    if not html_content: return

    try:
        lista_de_dataframes = pd.read_html(html_content, encoding='utf-8')
        if not lista_de_dataframes: return
        
        df_final = None
        if nome_base_arquivo.startswith("Relatorio_Sae_Solicitacao"):
            linha_cabecalho_original = 6
            print("   -> Relatório SAE detectado. Processando...")
            df_bruto_correto = None
            for df_candidata in lista_de_dataframes:
                if len(df_candidata) > linha_cabecalho_original:
                    df_bruto_correto = df_candidata
                    break 
            if df_bruto_correto is None: return

            linha_inicio_dados = linha_cabecalho_original + 1
            df_dados_puros = df_bruto_correto.iloc[linha_inicio_dados:, 1:].copy()
            cabecalho_final = ['Solicitação', 'Área', 'Assunto', 'Polo', 'Status', 'Situação', 'Data Solicitação', 'Data Previsão', 'Encaminhado Por']
            num_colunas = len(cabecalho_final)
            df_final_temp = df_dados_puros.iloc[:, :num_colunas]
            df_final_temp.columns = cabecalho_final
            df_final = df_final_temp
            
            # --- LINHA CRÍTICA ADICIONADA AQUI ---
            # Tenta converter a data para um formato limpo. Se falhar, deixa como está.
            df_final['Data Solicitação'] = pd.to_datetime(df_final['Data Solicitação'], errors='coerce', dayfirst=True)

        else:
            print(f"   -> Relatório {nome_base_arquivo} detectado. Corrigindo cabeçalho...")
            df_bruto = lista_de_dataframes[0]
            if not df_bruto.empty:
                novo_cabecalho = df_bruto.iloc[0]
                df_final = df_bruto[1:].copy()
                df_final.columns = novo_cabecalho
            else:
                df_final = df_bruto

        if df_final is not None:
            df_final.reset_index(drop=True, inplace=True)
            os.makedirs(diretorio_saida, exist_ok=True)
            nome_arquivo_saida = os.path.splitext(nome_base_arquivo)[0] + '.xlsx'
            caminho_completo_saida = os.path.join(diretorio_saida, nome_arquivo_saida)
            df_final.to_excel(caminho_completo_saida, index=False, engine='openpyxl')
            print(f"   -> Convertido com sucesso para '{caminho_completo_saida}'.")

    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar '{nome_base_arquivo}': {e}")

def executar_conversao():
    print("--- INICIANDO SCRIPT DE CONVERSÃO ---")
    try:
        if os.path.exists(DIRETORIO_SAIDA):
            for f in os.listdir(DIRETORIO_SAIDA): os.remove(os.path.join(DIRETORIO_SAIDA, f))
        arquivos_para_converter = [f for f in os.listdir(DIRETORIO_ENTRADA) if f.lower().endswith('.xls')]
    except FileNotFoundError:
        print(f"ERRO: A pasta de entrada '{DIRETORIO_ENTRADA}' não foi encontrada.")
        return False

    if not arquivos_para_converter: print(f"Nenhum arquivo .xls encontrado.")
    else:
        for arquivo in tqdm(arquivos_para_converter, desc="Convertendo relatórios"):
            converter_arquivo(os.path.join(DIRETORIO_ENTRADA, arquivo), DIRETORIO_SAIDA)
            
    print("\n--- PROCESSO DE CONVERSÃO CONCLUÍDO ---")
    return True

if __name__ == '__main__':
    if executar_conversao():
        volumetria.main()
    else:
        print("\nO script de volumetria não será executado devido a um erro na conversão.")
    input("\nPressione ENTER para fechar.")