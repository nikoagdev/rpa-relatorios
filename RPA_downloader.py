import pyautogui
import time
import subprocess
import sys
import os
import shutil
import glob
from dotenv import load_dotenv

# --- CONFIGURAÇÕES DINÂMICAS E SEGURAS ---

# Carrega as variáveis do arquivo .env (PORTAL_USER, PORTAL_PASSWORD)
load_dotenv()

# Esta é a mágica: pega o caminho absoluto da pasta onde o script está rodando.
# Agora, não importa onde você mova a pasta do projeto, os caminhos sempre funcionarão.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Caminhos relativos baseados no BASE_DIR
PASTA_ASSETS = os.path.join(BASE_DIR, 'assets')
PASTA_DESTINO_DOWNLOADS = os.path.join(BASE_DIR, 'relatorios_originais')

# Credenciais seguras carregadas do arquivo .env
MEU_USUARIO = os.getenv("PORTAL_USER")
MINHA_SENHA = os.getenv("PORTAL_PASSWORD")

# Caminho para o executável do portal (pode permanecer absoluto se for um app instalado)
CAMINHO_PORTAL = r"C:\Users\nikolas.alexandre\AppData\Local\Unicesumar\PortalAPP\portal-app-win32-x64\portal-app.exe"
PASTA_DOWNLOADS_WINDOWS = os.path.join(os.path.expanduser('~'), 'Downloads')


# --- FUNÇÕES AUXILIARES ATUALIZADAS ---

def encontrar_e_clicar(nome_imagem, tentativas=10, intervalo=1, confianca=0.9):
    """Procura por uma imagem na pasta 'assets' e clica nela."""
    caminho_imagem = os.path.join(PASTA_ASSETS, nome_imagem) # <-- MUDANÇA AQUI
    print(f"Procurando por '{nome_imagem}' com confiança de {confianca*100}%...")
    for _ in range(tentativas):
        try:
            posicao = pyautogui.locateCenterOnScreen(caminho_imagem, confidence=confianca)
            if posicao:
                pyautogui.click(posicao)
                print(f" -> Encontrado e clicado em '{nome_imagem}'.")
                return True
        except pyautogui.ImageNotFoundException:
            pass 
        time.sleep(intervalo)
    print(f" !!! ERRO: Imagem '{nome_imagem}' não encontrada na tela.")
    return False

def mover_e_renomear_ultimo_download(nome_final_arquivo):
    """Encontra o último arquivo .xls baixado e o move para a pasta do projeto."""
    print(f"Movendo o arquivo baixado para '{nome_final_arquivo}'...")
    try:
        time.sleep(3) 
        lista_de_arquivos = glob.glob(os.path.join(PASTA_DOWNLOADS_WINDOWS, '*.xls'))
        if not lista_de_arquivos:
            print(" !!! ERRO: Nenhum arquivo .xls encontrado na pasta de Downloads do Windows.")
            return False
        arquivo_mais_recente = max(lista_de_arquivos, key=os.path.getctime)
        print(f" -> Arquivo mais recente encontrado: '{os.path.basename(arquivo_mais_recente)}'")
        
        # O destino agora é a nossa pasta de projeto
        caminho_destino = os.path.join(PASTA_DESTINO_DOWNLOADS, nome_final_arquivo) # <-- MUDANÇA AQUI
        
        if os.path.exists(caminho_destino):
            os.remove(caminho_destino)
            print(f" -> Arquivo antigo '{nome_final_arquivo}' removido do destino.")
        shutil.move(arquivo_mais_recente, caminho_destino)
        print(f" -> Arquivo movido e renomeado com sucesso para '{caminho_destino}'")
        return True
    except Exception as e:
        print(f" !!! ERRO ao mover/renomear o arquivo: {e}")
        return False

# --- FUNÇÕES DE FLUXO ---

def fazer_login():
    """Realiza o login no portal e fecha pop-ups iniciais."""
    print("\n--- Iniciando processo de login ---")
    if not encontrar_e_clicar('usuario.png'): return False
    pyautogui.write(MEU_USUARIO, interval=0.05)
    if not encontrar_e_clicar('senha.png'): return False
    pyautogui.write(MINHA_SENHA, interval=0.05)
    if not encontrar_e_clicar('entrar.png'): return False
    print(" -> Login realizado com sucesso! Aguardando página principal...")
    time.sleep(15) 
    encontrar_e_clicar('ok_atencao.png', tentativas=5)
    time.sleep(2)
    return True

def processar_um_relatorio_sae(nome_final, imagem_fila):
    """Executa o ciclo completo para relatórios SAE: navega, filtra, gera, fecha e move."""
    print(f"\n--- Iniciando navegação para o relatório SAE: {nome_final} ---")
    
    if not (encontrar_e_clicar('relatorios_selecionados.png', tentativas=2, intervalo=0.5) or encontrar_e_clicar('relatorios_menu.png', tentativas=2, intervalo=0.5)):
        print("!!! ERRO: Botão 'Relatórios' não encontrado.")
        return False
    time.sleep(2)
    encontrar_e_clicar('ok_atencao.png', tentativas=3) 
    time.sleep(3)
    
    try:
        print(" -> Expandindo menus de navegação...")
        pos_comunicacao = pyautogui.locateCenterOnScreen('comunicacao_pasta.png', confidence=0.9)
        if not pos_comunicacao: raise Exception("Imagem 'comunicacao_pasta.png' não encontrada.")
        pyautogui.click(pos_comunicacao.x - 60, pos_comunicacao.y)
        time.sleep(1)

        pos_sae = pyautogui.locateCenterOnScreen('sae_pasta.png', confidence=0.9)
        if not pos_sae: raise Exception("Imagem 'sae_pasta.png' não encontrada.")
        pyautogui.click(pos_sae.x - 40, pos_sae.y)
        time.sleep(2)
    except Exception as e:
        print(f"!!! ERRO ao expandir menus: {e}")
        return False

    if not encontrar_e_clicar('solicitacoes_item.png'): return False
    print(" -> Navegação concluída. Abrindo janela de filtros...")
    time.sleep(3)

    try:
        print("Preenchendo filtros...")
        pos_fila = pyautogui.locateCenterOnScreen('fila_label.png', confidence=0.9)
        if not pos_fila: raise Exception("Label 'Fila' não encontrado")
        pyautogui.click(pos_fila.x + 200, pos_fila.y)
        time.sleep(1)
        if not encontrar_e_clicar(imagem_fila): return False

        pos_situacao = pyautogui.locateCenterOnScreen('situacao_label.png', confidence=0.9)
        if not pos_situacao: raise Exception("Label 'Situação' não encontrado")
        pyautogui.click(pos_situacao.x + 200, pos_situacao.y)
        time.sleep(1)
        if not encontrar_e_clicar('encaminhada_opcao.png'): return False
        
        pos_data = pyautogui.locateCenterOnScreen('data_solicitacao_label.png', confidence=0.9)
        if not pos_data: raise Exception("Label 'Data Solicitação' não encontrado")

        print(" -> Selecionando data de início...")
        pyautogui.click(pos_data.x + 100, pos_data.y)
        time.sleep(1)
        
        janeiro_encontrado = False
        for i in range(40):
            try:
                pyautogui.locateOnScreen('mes_janeiro_2025.png', confidence=0.9)
                janeiro_encontrado = True
                print(f" -> Mês 'Janeiro 2025' encontrado na tentativa {i+1}.")
                break
            except pyautogui.ImageNotFoundException:
                if not encontrar_e_clicar('seta_esquerda_calendario.png', tentativas=1, intervalo=0.1, confianca=0.95):
                    raise Exception("Não foi possível clicar na seta do calendário para voltar.")
                time.sleep(0.5)

        if not janeiro_encontrado: raise Exception("Não foi possível navegar até Janeiro de 2025.")
        if not encontrar_e_clicar('dia_1_calendario.png'): raise Exception("Não foi possível clicar no dia '1'.")

        print(" -> Selecionando data final...")
        pyautogui.click(pos_data.x + 250, pos_data.y)
        time.sleep(1)
        if not encontrar_e_clicar('dia_atual_calendario.png', confianca=0.7): 
            raise Exception("Não foi possível encontrar o dia atual no calendário.")
    except Exception as e:
        print(f"!!! ERRO ao preencher filtros: {e}")
        return False
    
    if not encontrar_e_clicar('gerar_excel.png'): return False
    if not encontrar_e_clicar('popup_cancelar.png', tentativas=60, confianca=0.95):
        print("!!! ERRO: Pop-up de download não encontrado.")
        return False
    
    if not mover_e_renomear_ultimo_download(nome_final): return False
    return True

def processar_relatorios_coordenacao():
    """Executa o fluxo de extração para os relatórios de Coordenação na tela principal."""
    print("\n\n--- INICIANDO GERAÇÃO DOS RELATÓRIOS DE COORDENAÇÃO ---")

    relatorios = [
        {"nome_final": "COORDENAÇÃO - ALUNOS.xls", "imagem_fila": "alunos_fila_opcao.png", "scroll_area": 1},
        {"nome_final": "COORDENAÇÃO - POLOS.xls", "imagem_fila": "polos_fila_opcao.png", "scroll_area": 2}
    ]
    try:
        print(" -> Definindo Campo para 'Solicitação'...")
        pos_campo = pyautogui.locateCenterOnScreen('campo_label.png', confidence=0.9)
        if not pos_campo: raise Exception("Label 'Campo' não encontrado")
        pyautogui.click(pos_campo.x + 150, pos_campo.y)
        time.sleep(1)
        if not encontrar_e_clicar('solicitacao_campo_opcao.png'): return False
        time.sleep(1)

        for index, relatorio in enumerate(relatorios):
            print(f"\n--- Processando: {relatorio['nome_final']} ---")
            
            pos_fila = pyautogui.locateCenterOnScreen('fila_label.png', confidence=0.9)
            if not pos_fila: raise Exception("Label 'Fila' não encontrado")
            pyautogui.click(pos_fila.x + 150, pos_fila.y)
            time.sleep(1)
            if not encontrar_e_clicar(relatorio['imagem_fila']): return False
            time.sleep(1)

            pos_area = pyautogui.locateCenterOnScreen('area_label.png', confidence=0.9)
            if not pos_area: raise Exception("Label 'Área' não encontrado")
            pyautogui.click(pos_area.x + 150, pos_area.y)
            time.sleep(1)
            pyautogui.move(0, 50)
            time.sleep(0.5)
            for _ in range(relatorio['scroll_area']):
                pyautogui.scroll(-100)
                time.sleep(0.5)
            if not encontrar_e_clicar('coordenacao_area_opcao.png'): return False
            time.sleep(1)

            if index == 0:
                print(" -> Selecionando Status 'NOVA' (primeira execução)...")
                pos_status = pyautogui.locateCenterOnScreen('status_label.png', confidence=0.9)
                if not pos_status: raise Exception("Label 'Status' não encontrado")
                pyautogui.click(pos_status.x + 150, pos_status.y)
                time.sleep(1)
                pyautogui.move(0, 50)
                time.sleep(0.5)
                pyautogui.scroll(-100)
                time.sleep(1)
                if not encontrar_e_clicar('nova_status_opcao.png'): return False
                time.sleep(1)
            else:
                print(" -> Status 'NOVA' já está selecionado. Pulando esta etapa.")
            
            if not encontrar_e_clicar('filtrar_botao.png'): return False
            print(" -> Filtrando dados. Aguardando 10 segundos...")
            time.sleep(10)

            try:
                pyautogui.locateOnScreen('registros_zero.png', confidence=0.95)
                print(f" -> 0 registros encontrados para '{relatorio['nome_final']}'. Pulando exportação.")
                encontrar_e_clicar('retirar_botao.png')
                time.sleep(3)
                continue
            except pyautogui.ImageNotFoundException:
                print(" -> Registros encontrados. Prosseguindo com a exportação.")

            pos_excel = pyautogui.locateCenterOnScreen('excel_botao.png', confidence=0.9)
            if not pos_excel: raise Exception("Botão 'Excel' não encontrado")
            pyautogui.click(pos_excel.x + 25, pos_excel.y)
            time.sleep(1)
            if not encontrar_e_clicar('sintetico_opcao.png'): return False
            time.sleep(5)

            nome_do_arquivo = relatorio['nome_final']
            print(f" -> Digitando nome do arquivo: {nome_do_arquivo}")
            pyautogui.write(nome_do_arquivo, interval=0.02)
            time.sleep(1)
            pyautogui.press('enter')
            print(" -> Arquivo salvo. Aguardando retorno à tela de filtros...")
            time.sleep(10)
    except Exception as e:
        print(f"!!! ERRO durante o processamento dos relatórios de Coordenação: {e}")
        return False
    return True

if __name__ == "__main__":
    print("=======================================================")
    print("     INICIANDO AUTOMAÇÃO DE DOWNLOAD DE RELATÓRIOS")
    print("=======================================================")
    
    rpa_sucesso = False
    try:
        if not os.path.exists(CAMINHO_PORTAL):
             raise FileNotFoundError(f"O executável do portal não foi encontrado em: {CAMINHO_PORTAL}")
        subprocess.Popen(CAMINHO_PORTAL)
        time.sleep(15)
        
        if not fazer_login():
            raise Exception("Erro durante o login.")
        
        if not processar_um_relatorio_sae("Relatorio_Sae_Solicitacao - Alunos.xls", "alunos_opcao.png"):
            print("!!! AVISO: FALHA na extração do relatório de Alunos SAE.")
        else:
            print(" -> Sucesso na extração do relatório de Alunos SAE.")
        
        time.sleep(5)
        
        if not processar_um_relatorio_sae("Relatorio_Sae_Solicitacao - Polos.xls", "polos_opcao.png"):
            print("!!! AVISO: FALHA na extração do relatório de Polos SAE.")
        else:
            print(" -> Sucesso na extração do relatório de Polos SAE.")
        
        print("\n\n--- EXTRAÇÃO DOS RELATÓRIOS SAE CONCLUÍDA ---")
        time.sleep(5)

        if not processar_relatorios_coordenacao():
            print("!!! AVISO: FALHA na extração dos relatórios de Coordenação.")
        else:
            print(" -> Sucesso na extração dos relatórios de Coordenação.")
        
        print("\n\n--- EXTRAÇÃO DE TODOS OS RELATÓRIOS CONCLUÍDA COM SUCESSO ---")
        rpa_sucesso = True

    except Exception as e:
        print(f"\nO processo de download foi interrompido por um erro: {e}")
    finally:
        pyautogui.hotkey('alt', 'f4')

    if rpa_sucesso:
        print("\n=======================================================")
        print("     ACIONANDO SCRIPT DE CONVERSÃO E VOLUMETRIA")
        print("=======================================================")
        
        try:
            # Agora os scripts estão na mesma pasta, a chamada é mais simples
            print("Executando converter.py...")
            subprocess.run([sys.executable, 'converter.py'], check=True)
            
            print("\nExecutando volumetria.py...")
            subprocess.run([sys.executable, 'volumetria.py'], check=True)
            
        except Exception as e:
            print(f"!!! ERRO ao executar os scripts de processamento: {e}")
    else:
        print("\nO processamento de arquivos não foi iniciado devido a uma falha no download.")

    print("\n\n--- PROCESSO DE AUTOMAÇÃO CONCLUÍDO ---")
    input("\nPressione ENTER para fechar.")

