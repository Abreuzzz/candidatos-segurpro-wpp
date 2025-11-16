"""Automação de envio de mensagens no WhatsApp Web.

Este módulo reorganiza o script original em funções reutilizáveis, permite a
configuração por linha de comando e adiciona otimizações para reduzir o tempo de
preparo entre execuções (por exemplo, reutilizar um perfil de navegador para
evitar leituras repetidas de QRCode quando desejado).
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# chromedriver-py garante que o binário esteja disponível mesmo em ambientes
# sem Chrome pré-instalado.
import chromedriver_py


DEFAULT_MESSAGE = (
    "Olá, boa noite! Tudo bem?\n"
    "Me chamo Karina e falo do Recrutamento e Seleção do Grupo Segurpro.\n"
    "Estamos com oportunidades para atuar como Vigilante Intermitente em uma "
    "operação especial de grande porte em São Paulo nos dias 12, 13 e 14/09.\n"
    "Você possui interesse?\n"
    "Se tiver disponibilidade imediata e reciclagem em dia, pedimos que compareça "
    "com seus documentos em:\n"
    "RUA DOS ITALIANOS, 988 - BOM RETIRO\n"
    "Amanhã (11/09) ÀS 8H.\n"
    "Contamos com sua presença!\n"
)


@dataclass
class RunStats:
    """Agrega os números enviados com sucesso e com falha."""

    successful_numbers: list[str]
    failed_numbers: list[str]

    @property
    def total(self) -> int:
        return len(self.successful_numbers) + len(self.failed_numbers)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Disparo de mensagens via WhatsApp Web")
    parser.add_argument(
        "--workbook",
        default="Pasta3.xlsx",
        help="Caminho para a planilha com os números (padrão: %(default)s)",
    )
    parser.add_argument(
        "--sheet",
        default="Planilha78",
        help="Nome da planilha a ser lida (padrão: %(default)s)",
    )
    parser.add_argument(
        "--column",
        default="A",
        help="Coluna com os números de telefone (padrão: %(default)s)",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Quantidade máxima de telefones a serem processados",
    )
    parser.add_argument(
        "--message-file",
        type=Path,
        help="Arquivo de texto com a mensagem a ser enviada",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=5.0,
        help="Tempo (s) de espera entre envios (padrão: %(default).1f)s",
    )
    parser.add_argument(
        "--report-number",
        dest="report_numbers",
        action="append",
        default=["11969257920", "11953075442"],
        help=(
            "Número que receberá o relatório final. Use várias vezes para múltiplos "
            "destinatários (padrão: %(default)s)"
        ),
    )
    parser.add_argument(
        "--user-data-dir",
        type=Path,
        help=(
            "Diretório de perfil do Chrome a ser reutilizado. Útil para reduzir a "
            "necessidade de ler o QRCode em execuções subsequentes."
        ),
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Roda o Chrome em modo headless (pode exigir sessão previamente autenticada).",
    )
    parser.add_argument(
        "--max-wait",
        type=int,
        default=60,
        help="Tempo máximo (s) de espera pelos elementos da página",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Nível de log (padrão: %(default)s)",
    )
    return parser.parse_args(argv)


def configure_logging(level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level),
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def load_numbers(workbook_path: str | Path, sheet_name: str, column: str, limit: int | None) -> list[str]:
    workbook = load_workbook(workbook_path, data_only=True)
    worksheet = workbook[sheet_name]
    numbers: list[str] = []

    for cell in worksheet[column]:
        if isinstance(cell.value, str) and cell.value.strip().lower() == "telefones":
            continue
        digits = normalize_phone(cell.value)
        if not digits:
            continue
        numbers.append(digits)
        if limit and len(numbers) >= limit:
            break

    workbook.close()
    # Remove duplicados mantendo a ordem original.
    seen = set()
    ordered_unique: list[str] = []
    for number in numbers:
        if number not in seen:
            ordered_unique.append(number)
            seen.add(number)
    return ordered_unique


def normalize_phone(value: object) -> str | None:
    if value is None:
        return None
    digits = ''.join(filter(str.isdigit, str(value)))
    return digits or None


def build_browser(max_wait: int, *, user_data_dir: Path | None, headless: bool) -> tuple[webdriver.Chrome, WebDriverWait]:
    options = webdriver.ChromeOptions()
    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    service = Service(executable_path=chromedriver_py.binary_path)
    browser = webdriver.Chrome(service=service, options=options)
    browser.maximize_window()
    wait = WebDriverWait(browser, max_wait)
    return browser, wait


def wait_for_message_box(wait: WebDriverWait):
    return wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '/html/body/div[1]/div/div[1]/div[3]/div/div[4]/div/footer/div[1]/div/'
                'span/div/div[2]/div/div[3]/div[1]/p',
            )
        )
    )


def send_message(browser: webdriver.Chrome, wait: WebDriverWait, number: str, message: str, delay: float) -> bool:
    try:
        browser.get(
            f'https://web.whatsapp.com/send/?phone=55{number}&text&type=phone_number&app_absent=0'
        )
        wait.until(EC.title_contains('WhatsApp'))
        input_box = wait_for_message_box(wait)
        input_box.send_keys(message)
        logging.info("Mensagem enviada para %s", number)
        return True
    except Exception as exc:  # noqa: BLE001 - queremos registrar qualquer falha.
        logging.error("Erro com %s: %s", number, exc)
        return False
    finally:
        time.sleep(delay)
        dismiss_alert(browser)


def dismiss_alert(browser: webdriver.Chrome) -> None:
    try:
        alert = browser.switch_to.alert
        alert.accept()
        logging.debug("Alert aceito")
    except Exception:  # noqa: BLE001 - ignoramos ausência de alertas.
        pass


def process_numbers(
    browser: webdriver.Chrome,
    wait: WebDriverWait,
    numbers: Iterable[str],
    message: str,
    delay: float,
) -> RunStats:
    successful: list[str] = []
    failed: list[str] = []

    for number in numbers:
        if send_message(browser, wait, number, message, delay):
            successful.append(number)
        else:
            failed.append(number)

    return RunStats(successful_numbers=successful, failed_numbers=failed)


def send_report(browser: webdriver.Chrome, wait: WebDriverWait, numbers: Iterable[str], message: str, delay: float) -> None:
    for number in numbers:
        normalized = normalize_phone(number)
        if not normalized:
            logging.warning("Número de relatório inválido: %s", number)
            continue
        send_message(browser, wait, normalized, message, delay)


def format_duration(seconds: float) -> str:
    return time.strftime("%H:%M:%S", time.gmtime(seconds))


def build_report(stats: RunStats, elapsed: float) -> str:
    tempo_medio = elapsed / stats.total if stats.total else 0
    return (
        f"Total de números: {stats.total}\n"
        f"Números enviados com sucesso: {len(stats.successful_numbers)}\n"
        f"Números enviados com erro: {len(stats.failed_numbers)}\n"
        f"Números com erro: {stats.failed_numbers}\n"
        f"Tempo de execução: {format_duration(elapsed)}\n"
        f"Tempo médio por número: {format_duration(tempo_medio)}\n"
    )


def load_message(message_file: Path | None) -> str:
    if not message_file:
        return DEFAULT_MESSAGE
    return message_file.read_text(encoding="utf-8").strip()


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    configure_logging(args.log_level)

    message = load_message(args.message_file)
    numbers = load_numbers(args.workbook, args.sheet, args.column, args.limit)
    if not numbers:
        logging.warning("Nenhum número válido encontrado na planilha.")
        return 1

    logging.info("Total de números únicos a enviar: %s", len(numbers))

    browser, wait = build_browser(
        args.max_wait,
        user_data_dir=args.user_data_dir,
        headless=args.headless,
    )

    start = time.time()
    try:
        stats = process_numbers(browser, wait, numbers, message, args.delay)
        elapsed = time.time() - start
        report_message = build_report(stats, elapsed)
        logging.info("Resumo:\n%s", report_message)
        send_report(browser, wait, args.report_numbers, report_message, args.delay)
        print(report_message)
    finally:
        browser.quit()

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

