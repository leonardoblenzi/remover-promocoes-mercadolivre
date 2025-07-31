import os
import sys
import time
import psutil
import pandas as pd
import subprocess
import requests
import socket
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (NoSuchElementException,
                                        WebDriverException,
                                        TimeoutException)
from urllib3.exceptions import ReadTimeoutError

# ===== CONFIGURAÇÕES =====
ARQUIVO_PLANILHA = r"C:\Users\USER\PycharmProjects\Remove Promo\DATA\Correcao Preco Diplany.xlsx"
COLUNA_MLB = "CODIGO_ANUNCIO"
TEMPO_ESPERA = 5
TEMPO_ESPERA_EXTRA = 3
TENTATIVAS_MAX = 3
TIMEOUT_PAGINA = 60  # segundos
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
    processos_mortos = 0
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] and any(x in proc.info['name'].lower() for x in ['msedge', 'edge']):
                proc.kill()
                processos_mortos += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    print(f"✔️ {processos_mortos} processos encerrados")
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

        # Aguarda até 30 segundos pela porta de depuração
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
        driver.set_page_load_timeout(TIMEOUT_PAGINA)
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


def reconectar_driver(driver):
    """Reinicia completamente o navegador e driver"""
    print("🔄 Reconectando o navegador...")
    try:
        if driver:
            driver.quit()
    except:
        pass

    matar_processos_edge()
    if not iniciar_edge_com_debug():
        raise Exception("Falha ao reiniciar o navegador")

    new_driver = conectar_selenium()
    if not new_driver:
        raise Exception("Falha ao reconectar o Selenium")

    return new_driver


def encontrar_botoes_promocao(driver):
    """Localiza todos os botões relevantes para remoção de promoções"""
    try:
        return WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH,
                                                 "//button[contains(., 'Deixar de participar') or "
                                                 "contains(., 'Alterar') or "
                                                 "contains(., 'Remover promoção')]")))
    except TimeoutException:
        return []


def processar_promocao(driver, botao, mlb):
    """Processa um botão de promoção individual"""
    try:
        texto = botao.text.strip()
        driver.execute_script("arguments[0].scrollIntoView();", botao)
        driver.execute_script("arguments[0].click();", botao)
        time.sleep(TEMPO_ESPERA)

        if "Deixar de participar" in texto:
            # Confirmar no popup
            try:
                confirmar = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH,
                                                "//button[contains(., 'Confirmar') or "
                                                "contains(., 'Deixar de participar') or "
                                                "contains(., 'Sim')]")))
                confirmar.click()
                time.sleep(TEMPO_ESPERA)
                return True
            except:
                return False

        elif "Alterar" in texto or "Remover promoção" in texto:
            # Clicar no botão "Deixar de participar" no popup
            try:
                deixar = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH,
                                                "//button[contains(., 'Deixar de participar')]")))
                deixar.click()
                time.sleep(TEMPO_ESPERA)

                # Confirmar no segundo popup
                try:
                    confirmar = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH,
                                                    "//button[contains(., 'Confirmar') or "
                                                    "contains(., 'Deixar de participar')]")))
                    confirmar.click()
                    time.sleep(TEMPO_ESPERA)
                    return True
                except:
                    return False
            except:
                return False
    except:
        return False


def remover_promocoes(driver, mlb):
    """Remove todas as promoções de um anúncio com tratamento robusto"""
    url = f"https://www.mercadolivre.com.br/anuncios/lista/promos/?search={mlb}"
    promocoes_removidas = 0

    for tentativa in range(1, TENTATIVAS_MAX + 1):
        try:
            try:
                driver.get(url)
                time.sleep(TEMPO_ESPERA)
            except (socket.timeout, ReadTimeoutError, WebDriverException) as e:
                print(f"⚠️ Timeout ao carregar página (Tentativa {tentativa}): {str(e)[:100]}")
                if tentativa == TENTATIVAS_MAX:
                    raise
                continue

            while True:
                botoes = encontrar_botoes_promocao(driver)
                if not botoes:
                    break

                for botao in botoes:
                    if processar_promocao(driver, botao, mlb):
                        promocoes_removidas += 1
                        print(f"✔️ Promoção removida (Tentativa {tentativa}) - {mlb}")
                    else:
                        print(f"⚠️ Falha ao processar botão (Tentativa {tentativa}) - {mlb}")

                # Verifica se ainda há promoções
                driver.get(url)
                time.sleep(TEMPO_ESPERA)

            return promocoes_removidas

        except Exception as e:
            print(f"❌ Erro ao processar {mlb} (Tentativa {tentativa}): {str(e)[:200]}")
            if tentativa < TENTATIVAS_MAX:
                try:
                    driver = reconectar_driver(driver)
                except Exception as reconect_error:
                    print(f"❌ Falha ao reconectar: {reconect_error}")
                    break
            else:
                driver.save_screenshot(f"erro_mlb_{mlb}.png")

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

        try:
            removidas = remover_promocoes(driver, mlb)

            if removidas == 0:
                falhas.append(mlb)
                total_sem_promocao += 1
                print(f"⚠️ Nenhuma promoção removida - {mlb}")
            else:
                total_promocoes_removidas += removidas
                print(f"✅ {removidas} promoção(ões) removida(s) - {mlb}")

            time.sleep(TEMPO_ESPERA_EXTRA)

        except Exception as e:
            print(f"❌ Erro crítico ao processar {mlb}: {str(e)[:200]}")
            falhas.append(mlb)
            try:
                driver = reconectar_driver(driver)
            except Exception as reconect_error:
                print(f"❌ Falha ao reconectar: {reconect_error}")
                break

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