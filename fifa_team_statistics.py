"""
Scrape FIFA team statistics with Selenium.

The page loads each statistics group dynamically after clicking category
buttons. This script clicks every configured category, waits for the FIFA
content container, extracts the rows, cleans the data, and creates one Excel
file per category.
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path
from typing import Any

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from convert_xlsx_to_csv import start_converter_process


DEFAULT_URL = (
    "https://www.fifa.com/en/tournaments/mens/worldcup/canadamexicousa2026/"
    "statistics/team-statistics"
)
MAIN_CONTENT_CLASS = "top-performer-group_mainContent__HbvU9"
MAIN_CONTENT_SELECTOR = f".{MAIN_CONTENT_CLASS}"
DEFAULT_CATEGORIES = [
    "Attacking",
    "Distribution",
    "Defending",
    "Discipline",
    "Goalkeeping",
    "Movement",
    "Physical",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Obtiene estadisticas de equipos de FIFA y genera Excel por categoria."
    )
    parser.add_argument("--url", default=DEFAULT_URL, help="URL de estadisticas FIFA.")
    parser.add_argument(
        "--out-dir",
        default="outputs/fifa_team_statistics",
        help="Carpeta donde se guardan los archivos Excel.",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Abre el navegador visible para depuracion.",
    )
    parser.add_argument(
        "--timeout-sec",
        type=int,
        default=45,
        help="Tiempo maximo de espera por elementos dinamicos.",
    )
    parser.add_argument(
        "--click-wait-sec",
        type=float,
        default=2.5,
        help="Pausa despues de cada click para permitir que carguen los datos.",
    )
    parser.add_argument(
        "--categories",
        nargs="*",
        default=DEFAULT_CATEGORIES,
        help="Categorias a consultar. Por defecto: todas las visibles principales.",
    )
    return parser.parse_args()


def normalize_header(value: str) -> str:
    value = re.sub(r"\s+", " ", str(value)).strip()
    value = value.replace("\u2195", "").replace("\u2191", "").replace("\u2193", "")
    return re.sub(r"\s+", " ", value).strip()


def normalize_team(value: Any) -> str:
    text = "" if pd.isna(value) else str(value)
    return re.sub(r"\s+", " ", text).strip()


def safe_metric_name(value: str) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def safe_file_name(value: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(value)).strip("_")
    return name or "category"


def has_any_data(series: pd.Series) -> bool:
    values = series.dropna()
    if values.empty:
        return False

    text_values = values.astype(str).str.strip()
    return text_values.ne("").any()


def to_number(value: Any) -> Any:
    if pd.isna(value):
        return pd.NA

    text = str(value).strip()
    if not text:
        return pd.NA

    text = text.replace(",", "")
    text = re.sub(r"[%+]", "", text)
    text = re.sub(r"[^\d.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return pd.NA

    try:
        number = float(text)
    except ValueError:
        return value

    return int(number) if number.is_integer() else number


def build_driver(headed: bool) -> WebDriver:
    options = Options()
    if not headed:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1600,1000")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=options)


def dismiss_banners(driver: WebDriver, timeout_sec: int) -> None:
    labels = [
        "Accept All",
        "Accept all",
        "I Accept",
        "Agree",
        "Aceptar",
        "Aceptar todo",
        "Allow all",
    ]
    wait = WebDriverWait(driver, min(timeout_sec, 5))
    for label in labels:
        try:
            button = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//button[contains(translate(., 'ACEPTLOWRI', 'aceptlowri'), '{label.lower()}')]")
                )
            )
            button.click()
            return
        except Exception:
            continue


def wait_for_main_content(driver: WebDriver, timeout_sec: int):
    return WebDriverWait(driver, timeout_sec).until(
        EC.presence_of_element_located((By.CLASS_NAME, MAIN_CONTENT_CLASS))
    )


def get_main_content_text(driver: WebDriver) -> str:
    try:
        return driver.find_element(By.CLASS_NAME, MAIN_CONTENT_CLASS).text
    except WebDriverException:
        return ""


def click_category(driver: WebDriver, category: str, timeout_sec: int, wait_sec: float) -> None:
    wait = WebDriverWait(driver, timeout_sec)
    previous_text = get_main_content_text(driver)
    xpath = (
        "//button[normalize-space()=$label] | "
        "//*[@role='button' and normalize-space()=$label] | "
        "//*[self::button or @role='button'][contains(normalize-space(), $label)]"
    )
    xpath = xpath.replace("$label", f"'{category}'")

    button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", button)
    wait_for_main_content(driver, timeout_sec)
    time.sleep(wait_sec)

    try:
        wait.until(lambda d: get_main_content_text(d) and get_main_content_text(d) != previous_text)
    except TimeoutException:
        pass


def load_all_rows(driver: WebDriver, timeout_sec: int) -> None:
    previous_text = ""
    stable_rounds = 0

    for _ in range(10):
        driver.execute_script("window.scrollBy(0, 900);")
        time.sleep(0.6)

        for label in ["Show more", "Load more", "View more", "Mostrar mas", "Ver mas"]:
            try:
                button = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, f"//button[contains(normalize-space(), '{label}')]")
                    )
                )
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
                break
            except Exception:
                continue

        current_text = get_main_content_text(driver)
        if current_text == previous_text:
            stable_rounds += 1
        else:
            stable_rounds = 0
        previous_text = current_text

        if stable_rounds >= 2:
            break

    driver.execute_script("window.scrollTo(0, 0);")
    wait_for_main_content(driver, timeout_sec)


def extract_rows_from_current_category(driver: WebDriver) -> list[dict[str, str]]:
    return driver.execute_script(
        """
        const clean = (text) => (text || '').replace(/\\s+/g, ' ').trim();
        const container = document.querySelector(arguments[0]);

        if (!container) {
          return [];
        }

        const headerSelectors = [
          '[role="columnheader"]',
          'thead th',
          '[class*="header"]',
          '[class*="Header"]'
        ].join(', ');
        const rowSelectors = [
          '[role="row"]',
          'tbody tr',
          '[class*="row"]',
          '[class*="Row"]'
        ].join(', ');
        const cellSelectors = [
          '[role="cell"]',
          '[role="gridcell"]',
          'td',
          'th',
          '[class*="cell"]',
          '[class*="Cell"]'
        ].join(', ');

        let headers = Array.from(container.querySelectorAll(headerSelectors))
          .map((node) => clean(node.innerText))
          .filter(Boolean);
        let rows = Array.from(container.querySelectorAll(rowSelectors))
          .filter((row) => row.querySelectorAll(cellSelectors).length > 1);

        if (headers.length && rows.length) {
          return rows.map((row) => {
            const cells = Array.from(row.querySelectorAll(cellSelectors))
              .map((node) => clean(node.innerText))
              .filter(Boolean);
            const item = {};
            headers.forEach((header, index) => { item[header] = cells[index] || ''; });
            return item;
          });
        }

        const lines = (container.innerText || '').split(/\\n+/).map(clean).filter(Boolean);
        const rankIndex = lines.findIndex((line) => /^Rank$/i.test(line));
        const teamIndex = lines.findIndex((line) => /^Team$/i.test(line));

        if (rankIndex === -1 || teamIndex === -1 || teamIndex <= rankIndex) {
          return [];
        }

        const firstRankAfterHeader = lines.findIndex(
          (line, index) => index > teamIndex && /^\\d+$/.test(line)
        );

        if (firstRankAfterHeader === -1) {
          return [];
        }

        headers = lines
          .slice(rankIndex, firstRankAfterHeader)
          .filter((line) => !/^Glossary$/i.test(line));
        const values = lines.slice(firstRankAfterHeader);
        const width = headers.length;
        const extracted = [];

        for (let index = 0; index + width <= values.length; index += width) {
          const cells = values.slice(index, index + width);
          if (!/^\\d+$/.test(cells[0])) {
            continue;
          }

          const item = {};
          headers.forEach((header, headerIndex) => {
            item[header] = cells[headerIndex] || '';
          });
          extracted.push(item);
        }

        return extracted;
        """,
        MAIN_CONTENT_SELECTOR,
    )


def scrape_fifa_categories(
    url: str,
    categories: list[str],
    headed: bool,
    timeout_sec: int,
    click_wait_sec: float,
) -> pd.DataFrame:
    driver = build_driver(headed)
    all_rows: list[dict[str, str]] = []

    try:
        driver.get(url)
        dismiss_banners(driver, timeout_sec)
        wait_for_main_content(driver, timeout_sec)

        for category in categories:
            print(f"Consultando categoria: {category}")
            try:
                click_category(driver, category, timeout_sec, click_wait_sec)
                load_all_rows(driver, timeout_sec)
                rows = extract_rows_from_current_category(driver)
            except TimeoutException:
                print(f"AVISO: no encontre o no cargo el boton {category}.")
                continue

            if not rows:
                print(f"AVISO: {category} cargo, pero no pude extraer filas.")
                continue

            for row in rows:
                row["Category"] = category
            all_rows.extend(rows)
    finally:
        driver.quit()

    if not all_rows:
        raise RuntimeError("No pude extraer filas de ninguna categoria.")

    return pd.DataFrame(all_rows)


def clean_category_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_header(column) for column in df.columns]
    df = df.loc[:, [column for column in df.columns if column]]

    if "Team" not in df.columns:
        possible_team_cols = [c for c in df.columns if "team" in c.lower()]
        if possible_team_cols:
            df = df.rename(columns={possible_team_cols[0]: "Team"})
        else:
            raise RuntimeError(f"No encontre la columna Team. Columnas: {list(df.columns)}")

    if "Category" not in df.columns:
        df["Category"] = "Statistics"

    df["Team"] = df["Team"].map(normalize_team)
    df["Category"] = df["Category"].map(safe_metric_name)
    df = df[df["Team"].ne("")]

    for column in df.columns:
        if column in {"Team", "Category"}:
            continue
        converted = df[column].map(to_number)
        present = converted.notna()
        numeric_ratio = pd.to_numeric(converted[present], errors="coerce").notna().mean()
        if present.any() and numeric_ratio >= 0.4:
            df[column] = pd.to_numeric(converted, errors="coerce")

    required_columns = {"Category", "Rank", "Team"}
    data_columns = [
        column
        for column in df.columns
        if column in required_columns or has_any_data(df[column])
    ]
    df = df[data_columns]

    columns = list(df.columns)
    ordered_columns = []
    for column in ["Category", "Rank", "Team"]:
        if column in columns:
            ordered_columns.append(column)
    ordered_columns.extend([column for column in columns if column not in ordered_columns])
    df = df[ordered_columns]

    if "Rank" in df.columns:
        df = df.sort_values(["Rank", "Team"], na_position="last")
    else:
        df = df.sort_values("Team")

    return df.reset_index(drop=True)


def autosize_excel_columns(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    for index, column in enumerate(df.columns, start=1):
        values = [str(column), *df[column].dropna().astype(str).head(100).tolist()]
        width = min(max(len(value) for value in values) + 2, 45)
        worksheet.column_dimensions[worksheet.cell(row=1, column=index).column_letter].width = width


def save_category_excels(raw: pd.DataFrame, out_dir: Path) -> list[Path]:
    excel_dir = out_dir / "excel_por_categoria"
    excel_dir.mkdir(parents=True, exist_ok=True)
    saved_files: list[Path] = []

    for category, category_df in raw.groupby("Category", sort=False):
        cleaned = clean_category_table(category_df)
        file_path = excel_dir / f"{safe_file_name(category)}.xlsx"

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            cleaned.to_excel(writer, sheet_name="Datos limpios", index=False)
            autosize_excel_columns(writer, "Datos limpios", cleaned)

            numeric_summary = cleaned.describe(include="number").transpose()
            if not numeric_summary.empty:
                numeric_summary.to_excel(writer, sheet_name="Resumen numerico")
                autosize_excel_columns(writer, "Resumen numerico", numeric_summary.reset_index())

        saved_files.append(file_path)

    return saved_files


def main() -> int:
    args = parse_args()
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    print("Abriendo pagina FIFA con Selenium y consultando categorias...")
    raw = scrape_fifa_categories(
        url=args.url,
        categories=args.categories,
        headed=args.headed,
        timeout_sec=args.timeout_sec,
        click_wait_sec=args.click_wait_sec,
    )
    print("Limpiando tablas y creando Excel por categoria...")
    saved_files = save_category_excels(raw, out_dir)

    print(f"Listo. Archivos guardados en: {out_dir.resolve()}")
    start_converter_process()
    for file_path in saved_files:
        print(f"- {file_path.resolve()}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())        
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
