# ====== IMPORTS (bibliotecas usadas) ======
import time
import sys
import pandas as pd
import schedule
from pathlib import Path
from urllib.parse import quote

# Imports do Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ====== CONFIGURAÇÕES DO USUÁRIO ======
ARQUIVO_EXCEL = "DataBase.xlsx"
HORA_ENVIO = "19:00"
PAUSA_ENTRE_CONTATOS = 1
PAUSA_INICIAL_LOGIN = 3
MANTER_JANELA_ABERTA = True

# ====== VARIÁVEIS GLOBAIS ======
driver = None

# ---------------- Funções utilitárias ----------------

def validar_planilha(df: pd.DataFrame) -> pd.DataFrame:
    colunas_esperadas = {"numero", "mensagem"}
    if not colunas_esperadas.issubset(df.columns.str.lower()):
        raise ValueError("A planilha precisa ter as colunas: 'numero' e 'mensagem'.")

    df = df.rename(columns={c: c.lower() for c in df.columns})
    df = df.dropna(how="all")
    df = df.dropna(subset=["numero", "mensagem"])
    df["numero"] = df["numero"].astype(str).str.replace(r"\D", "", regex=True)
    df["mensagem"] = df["mensagem"].astype(str).apply(lambda s: s.strip())
    df = df[df["numero"].str.len() > 0]

    if df.empty:
        raise ValueError("A planilha ficou vazia após validação. Verifique os dados.")
    return df


def iniciar_whatsapp():
    global driver
    session_dir = Path(__file__).with_name("whatsapp_session")
    session_dir.mkdir(exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={session_dir}")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    driver.get("https://web.whatsapp.com/")

    print("➡️  Se for o primeiro uso, escaneie o QR Code no WhatsApp Web.")
    time.sleep(PAUSA_INICIAL_LOGIN)

    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true']"))
        )
        print("✅ WhatsApp Web pronto.")
    except Exception:
        print("⚠️ Não consegui confirmar o carregamento. Tentando mesmo assim...")


def abrir_conversa_com_texto(numero: str, mensagem: str):
    msg_codificada = quote(mensagem)
    url = f"https://web.whatsapp.com/send?phone={numero}&text={msg_codificada}"
    driver.get(url)

    # Espera a caixa de texto ficar visível
    caixa = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true'][data-tab]"))
    )

    # Garante que o texto já está preenchido na caixa antes de enviar
    WebDriverWait(driver, 10).until(lambda d: mensagem in caixa.text)
    return caixa


def enviar_mensagem(driver, numero, mensagem):
    try:
        # Abre a conversa no WhatsApp Web
        url = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem}"
        driver.get(url)
        time.sleep(10)  # espera o carregamento da conversa

        # Localiza a caixa de texto
        campo_msg = driver.find_element("xpath", '//div[@contenteditable="true"][@data-tab="10"]')

        # Garante que a mensagem foi escrita e envia com ENTER
        campo_msg.send_keys(Keys.ENTER)

        print(f"✅ Mensagem enviada para {numero}: {mensagem}")
        time.sleep(5)  # espera para evitar duplicação
    except Exception as e:
        print(f"❌ Erro ao enviar para {numero}: {e}")

def enviar_de_planilha():
    caminho = Path(__file__).with_name(ARQUIVO_EXCEL)
    if not caminho.exists():
        print(f"❌ Arquivo Excel não encontrado: {caminho}")
        return

    try:
        df = pd.read_excel(caminho)
        df = validar_planilha(df)
    except Exception as e:
        print(f"❌ Erro ao ler/validar a planilha: {e}")
        return

    total = len(df)
    print(f"📄 Enviando {total} mensagem(ns) a partir de '{ARQUIVO_EXCEL}'...")

    for _, row in df.iterrows():
        numero = row["numero"]
        mensagem = row["mensagem"]
        enviar_mensagem (driver, numero, mensagem)
        time.sleep(PAUSA_ENTRE_CONTATOS)

    print("🎉 Finalizado o envio da planilha.")


def agendar_envio_diario():
    schedule.every().day.at(HORA_ENVIO).do(enviar_de_planilha)
    print(f"⏰ Agendado envio diário às {HORA_ENVIO}. Deixe esta janela aberta.")

    while True:
        schedule.run_pending()
        time.sleep(1)


def main():
    try:
        iniciar_whatsapp()
        enviar_de_planilha()
        agendar_envio_diario()
    except KeyboardInterrupt:
        print("\n🛑 Interrompido pelo usuário.")
    finally:
        if driver and not MANTER_JANELA_ABERTA:
            driver.quit()


if __name__ == "__main__":
    main()
