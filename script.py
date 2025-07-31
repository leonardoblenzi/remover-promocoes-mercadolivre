import os
import sys
import time
import psutil
import pandas as pd
import subprocess
import requests
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, WebDriverException

# ===== CONFIGURAÇÕES =====
ARQUIVO_PLANILHA = r"C:\Users\USER\PycharmProjects\Remove Promo\DATA\PREÇO CORRIGIR FRETE GRATIS SEM 12.xlsx"
COLUNA_MLB = "CODIGO_ANUNCIO"
TEMPO_ESPERA = 5
TEMPO_ESPERA_EXTRA = 3  # tempo adicional entre ações
TENTATIVAS_MAX = 3
ARQUIVO_FALHAS = "mlbs_falhados.txt"

EDGE_PROFILE_PATH = r"C:\Users\USER\AppData\Local\Microsoft\Edge\User Data"
EDGE_PROFILE_NAME = "Profile 1"
EDGE_DRIVER_PATH = r"C:\WebDriver\msedgedriver.exe"
EDGE_BINARY_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
DEBUG_PORT = 9222


# ===== FUNÇÕES =====
def verificar_porta_debug():
    """Verifica se a porta de depuração está respondendo"""
    try:
        response = requests.get(f"http://localhost:{DEBUG_PORT}/json/version", timeout=5)
        return response.status_code == 200
    except:
        return False


def matar_processos_edge():
    """Encerra todos os processos do Edge de forma robusta"""
    print("🔴 Encerrando processos do Edge...")
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] and any(x in proc.info['name'].lower() for x in ['msedge', 'edge']):
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    time.sleep(5)


def iniciar_edge_com_debug():
    """Inicia o Edge manualmente com porta de depuração"""
    print("🚀 Iniciando Edge com depuração remota...")
    cmd = [
        EDGE_BINARY_PATH,
        f"--user-data-dir={EDGE_PROFILE_PATH}",
        f"--profile-directory={EDGE_PROFILE_NAME}",
        f"--remote-debugging-port={DEBUG_PORT}",
        "--no-first-run",
        "about:blank"
    ]

    try:
        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        for _ in range(30):
            if verificar_porta_debug():
                return True
            time.sleep(1)

        print("❌ Timeout: Porta de depuração não respondeu")
        return False
    except Exception as e:
        print(f"❌ Falha ao iniciar Edge: {e}")
        return False


def conectar_selenium():
    """Conecta Selenium ao Edge usando a nova sintaxe"""
    options = Options()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")

    try:
        service = Service(executable_path=EDGE_DRIVER_PATH)
        driver = webdriver.Edge(service=service, options=options)
        driver.get("about:blank")  # Teste de conexão
        return driver
    except Exception as e:
        print(f"❌ Falha ao conectar Selenium: {e}")
        return None


def verificar_ambiente():
    """Verifica se todos os requisitos estão atendidos"""
    erros = []

    if not os.path.exists(EDGE_DRIVER_PATH):
        erros.append(f"EdgeDriver não encontrado em: {EDGE_DRIVER_PATH}")

    if not os.path.exists(EDGE_BINARY_PATH):
        erros.append(f"Executável do Edge não encontrado em: {EDGE_BINARY_PATH}")

    if not os.path.exists(ARQUIVO_PLANILHA):
        erros.append(f"Planilha não encontrada: {ARQUIVO_PLANILHA}")

    if erros:
        print("\n❌ Erros de configuração:")
        for erro in erros:
            print(f"- {erro}")
        print("\n🔍 Soluções:")
        print("- Baixe o EdgeDriver em: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
        print("- Verifique os caminhos configurados")
        sys.exit(1)


def remover_promocoes(driver, mlb):
    """
    Remove todas as promoções de um anúncio (diretas ou via popup).
    Retorna a quantidade de promoções removidas.
    """
    url = f"https://www.mercadolivre.com.br/anuncios/lista/promos/?search={mlb}"
    driver.get(url)
    time.sleep(TEMPO_ESPERA)

    promocoes_removidas = 0

    while True:
        promocoes_restantes = False
        try:
            acoes = driver.find_elements(By.XPATH, "//button[contains(., 'Deixar de participar') or contains(., 'Alterar')]")
            if not acoes:
                break

            for acao in acoes:
                promocoes_restantes = True
                texto = acao.text.strip()

                # Caso o botão direto "Deixar de participar" esteja visível
                if "Deixar de participar" in texto:
                    acao.click()
                    time.sleep(TEMPO_ESPERA)

                    # Confirmar no segundo popup (botão azul)
                    try:
                        confirmar_final = driver.find_element(By.XPATH, "//button[contains(., 'Deixar de participar')]")
                        confirmar_final.click()
                        promocoes_removidas += 1
                        print(f"✔️ Promoção removida diretamente - {mlb}")
                    except NoSuchElementException:
                        print(f"⚠️ Confirmação final não encontrada - {mlb}")

                    time.sleep(TEMPO_ESPERA_EXTRA)

                # Caso precise clicar em "Alterar"
                elif "Alterar" in texto:
                    acao.click()
                    time.sleep(TEMPO_ESPERA)

                    try:
                        deixar = driver.find_element(By.XPATH, "//button[contains(., 'Deixar de participar')]")
                        deixar.click()
                        time.sleep(TEMPO_ESPERA)

                        # Confirmar no segundo popup (botão azul)
                        try:
                            confirmar_final = driver.find_element(By.XPATH, "//button[contains(., 'Deixar de participar')]")
                            confirmar_final.click()
                            promocoes_removidas += 1
                            print(f"✔️ Promoção removida via popup - {mlb}")
                        except NoSuchElementException:
                            print(f"⚠️ Confirmação final não encontrada - {mlb}")

                    except NoSuchElementException:
                        print(f"⚠️ Botão 'Deixar de participar' não encontrado no popup - {mlb}")

                    time.sleep(TEMPO_ESPERA_EXTRA)

            # Atualiza a página para verificar se restam promoções
            driver.get(url)
            time.sleep(TEMPO_ESPERA)

        except Exception as e:
            print(f"❌ Erro ao processar {mlb}: {str(e)[:200]}")
            break

        if not promocoes_restantes:
            break

    return promocoes_removidas


# ===== EXECUÇÃO PRINCIPAL =====
def main():
    verificar_ambiente()
    matar_processos_edge()

    if not iniciar_edge_com_debug():
        sys.exit(1)

    print("🔗 Conectando Selenium...")
    driver = conectar_selenium()
    if not driver:
        sys.exit(1)

    print(f"✅ Edge conectado com sucesso via porta {DEBUG_PORT}")

    try:
        df = pd.read_excel(ARQUIVO_PLANILHA)
        mlbs = df[COLUNA_MLB].dropna().astype(str).str.strip().tolist()
        print(f"\n📊 Total de anúncios a processar: {len(mlbs)}")
    except Exception as e:
        print(f"❌ Erro ao carregar planilha: {e}")
        driver.quit()
        sys.exit(1)

    falhas = []
    total_promocoes_removidas = 0
    total_sem_promocao = 0

    for i, mlb in enumerate(mlbs, 1):
        mlb = mlb.strip()
        if not mlb.startswith('MLB'):
            print(f"⚠️ MLB inválido: {mlb} - Pulando...")
            continue

        print(f"\n🔎 Processando {i}/{len(mlbs)} - MLB: {mlb}")
        removidas = remover_promocoes(driver, mlb)

        if removidas == 0:
            falhas.append(mlb)
            total_sem_promocao += 1
            print(f"⚠️ Nenhuma promoção removida ou anúncio já estava sem promoções - {mlb}")
        else:
            total_promocoes_removidas += removidas
            print(f"✅ {removidas} promoção(ões) removida(s) do anúncio {mlb}")

        time.sleep(TEMPO_ESPERA_EXTRA)

    # Salva os mlbs que falharam
    if falhas:
        with open(ARQUIVO_FALHAS, "w", encoding="utf-8") as f:
            f.write("\n".join(falhas))
        print(f"\n⚠️ Alguns anúncios não tiveram promoções removidas. Lista salva em {ARQUIVO_FALHAS}")

    driver.quit()
    print("\n" + "=" * 50)
    print(f"✅ PROCESSO CONCLUÍDO")
    print(f"🔹 Total de anúncios processados: {len(mlbs)}")
    print(f"🔹 Total de promoções removidas: {total_promocoes_removidas}")
    print(f"🔹 Anúncios já sem promoções: {total_sem_promocao}")
    print(f"🔹 Anúncios com falha: {len(falhas)}")
    print("=" * 50)


if __name__ == "__main__":
    main()
