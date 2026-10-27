"""Envio em massa pelo WhatsApp Web (Selenium).

Uso:
  python enviar.py              # simulação: só lista para quem enviaria
  python enviar.py --teste      # envia só para a linha "Meu numero"
  python enviar.py --enviar     # envia de verdade para todas as escolas

Mensagem: edite mensagem.txt ({escola} é trocado pelo nome da escola).
Quem já recebeu fica registrado em enviados.csv e não recebe de novo.
"""
import argparse
import csv
import random
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

import openpyxl
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

BASE = Path(__file__).parent
PLANILHA = BASE / "escolas.xlsx"
MENSAGEM = BASE / "mensagem.txt"
LOG = BASE / "enviados.csv"
PERFIL = BASE / "perfil_chrome"  # guarda o login do WhatsApp (QR code só na 1ª vez)

VIDEO = BASE / "locateiq-escolas.mp4"  # enviado logo depois do texto
INTERVALO = (45, 60)  # segundos entre contatos (mínimo 45, com variação aleatória)
TESTE_NOME = "meu numero"

CAIXA_MSG = '//footer//div[@contenteditable="true"]'
ERRO_NUMERO = '//div[contains(text(),"inválido") or contains(text(),"invalid")]'


def normalizar(valor):
    """Devolve (numero_com_55, aviso)."""
    d = re.sub(r"\D", "", str(valor or ""))
    if d.startswith("55") and len(d) in (12, 13):
        d = d[2:]
    if len(d) not in (10, 11):
        return None, f"tamanho inesperado ({len(d)} dígitos)"
    aviso = None
    if len(d) == 10:
        aviso = "10 dígitos (provável telefone fixo ou falta o 9)"
    return "55" + d, aviso


def ler_contatos():
    ws = openpyxl.load_workbook(PLANILHA).active
    contatos = []
    for escola, zap, *_ in ws.iter_rows(min_row=3, values_only=True):
        if escola and zap:
            contatos.append((str(escola).strip(), zap))
    return contatos


def ja_enviados():
    if not LOG.exists():
        return set()
    with open(LOG, newline="", encoding="utf-8") as f:
        return {r["numero"] for r in csv.DictReader(f) if r["status"] == "ok"}


def registrar(escola, numero, status):
    novo = not LOG.exists()
    with open(LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if novo:
            w.writerow(["escola", "numero", "status", "hora"])
        w.writerow([escola, numero, status, time.strftime("%Y-%m-%d %H:%M:%S")])


def abrir_navegador():
    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={PERFIL}")
    driver = webdriver.Chrome(options=opts)
    driver.get("https://web.whatsapp.com")
    print("Se aparecer o QR code, escaneie com o celular...")
    WebDriverWait(driver, 180).until(
        EC.presence_of_element_located((By.XPATH, '//div[@id="pane-side"]'))
    )
    return driver


def enviar(driver, numero, texto):
    driver.get(f"https://web.whatsapp.com/send?phone={numero}&text={quote(texto)}")
    espera = WebDriverWait(driver, 40)
    try:
        espera.until(
            lambda d: d.find_elements(By.XPATH, CAIXA_MSG)
            or d.find_elements(By.XPATH, ERRO_NUMERO)
        )
    except TimeoutException:
        return "timeout"
    if driver.find_elements(By.XPATH, ERRO_NUMERO):
        return "numero_invalido"
    if driver.find_elements(By.XPATH, '//*[@aria-label="Remover anexo"]'):
        # sobra de execução interrompida: fecha a prévia de anexo pendente
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        espera.until(EC.presence_of_element_located((By.XPATH, CAIXA_MSG)))
    caixa = driver.find_element(By.XPATH, CAIXA_MSG)
    time.sleep(random.uniform(1.5, 3))
    caixa.send_keys(Keys.ENTER)
    time.sleep(random.uniform(3, 5))  # deixa o texto sair antes do vídeo
    if VIDEO.exists():
        return "ok" if enviar_video(driver, espera) else "texto_ok_video_falhou"
    return "ok"


def enviar_video(driver, espera):
    try:
        # o WhatsApp só cria o input de "Fotos e vídeos" ao clicar no item do menu;
        # capturamos esse input em vez de deixar abrir a janela de arquivos do Windows
        driver.execute_script(
            """const o = HTMLInputElement.prototype.click;
            HTMLInputElement.prototype.click = function () {
                if (this.type === 'file') { window.__fi = this; return; }
                return o.apply(this, arguments);
            };
            window.__fi = null;"""
        )
        espera.until(
            EC.element_to_be_clickable((By.XPATH, '//footer//button[@aria-label="Anexar"]'))
        ).click()
        espera.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@role="menuitem" or self::li][contains(.,"Fotos e vídeos")]'))
        ).click()
        campo = espera.until(lambda d: d.execute_script("return window.__fi"))
        driver.execute_script("document.body.appendChild(arguments[0]);", campo)
        campo.send_keys(str(VIDEO.resolve()))
        # o botão só fica ativo depois que o vídeo termina de carregar na prévia
        WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@aria-label="Enviar 1 item selecionado"]'))
        ).click()
        time.sleep(random.uniform(10, 14))  # tempo de upload antes de sair da conversa
        return True
    except TimeoutException:
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--enviar", action="store_true", help="envia de verdade")
    ap.add_argument("--teste", action="store_true", help="envia só para 'Meu numero'")
    args = ap.parse_args()

    template = MENSAGEM.read_text(encoding="utf-8").strip()
    contatos = ler_contatos()
    if args.teste:
        contatos = [c for c in contatos if c[0].lower() == TESTE_NOME]
    else:
        contatos = [c for c in contatos if c[0].lower() != TESTE_NOME]

    feitos = ja_enviados()
    fila = []
    for escola, zap in contatos:
        numero, aviso = normalizar(zap)
        if not numero:
            print(f"[PULADO] {escola}: {zap} -> {aviso}")
            continue
        if numero in feitos and not args.teste:
            print(f"[JÁ ENVIADO] {escola}: {numero}")
            continue
        if aviso:
            print(f"[ATENÇÃO] {escola}: {numero} -> {aviso}")
        fila.append((escola, numero))

    print(f"\n{len(fila)} mensagem(ns) na fila.")
    if not args.enviar and not args.teste:
        print("\nSIMULAÇÃO - nada foi enviado. Prévia da 1ª mensagem:\n")
        if fila:
            print(template.format(escola=fila[0][0]))
        for escola, numero in fila:
            print(f"  -> {escola} ({numero})")
        print("\nPara enviar: python enviar.py --enviar  (ou --teste primeiro)")
        return

    driver = abrir_navegador()
    try:
        for i, (escola, numero) in enumerate(fila, 1):
            texto = template.format(escola=escola)
            status = enviar(driver, numero, texto)
            registrar(escola, numero, status)
            print(f"[{i}/{len(fila)}] {escola} ({numero}): {status}")
            if i < len(fila):
                pausa = random.uniform(*INTERVALO)
                print(f"   aguardando {pausa:.0f}s...")
                time.sleep(pausa)
    except KeyboardInterrupt:
        print("\nInterrompido. O que já foi enviado está em enviados.csv.")
    finally:
        driver.quit()


if __name__ == "__main__":
    sys.exit(main())
