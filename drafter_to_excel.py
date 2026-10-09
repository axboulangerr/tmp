from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

MAX_DRAFTS = 50


def validate_url(url: str) -> None:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or parsed.hostname not in {
        "drafter.lol",
        "www.drafter.lol",
    } or not parsed.path.startswith("/draft/"):
        raise ValueError(f"Lien Drafter invalide : {url}")


def read_team_name(page, side: str) -> str:
    label = f"{side} side"
    element = page.get_by_text(label, exact=True).first

    for _ in range(6):
        text = re.sub(r"\s+", " ", element.inner_text()).strip()
        before_label, separator, _ = text.partition(label)
        if separator:
            name = re.sub(r"\s+(?:Ready|Not ready)\s*$", "", before_label, flags=re.I).strip()
            if name and name.casefold() != "ready":
                return name
        element = element.locator("xpath=..")

    return side.upper()


def champion_id_from_image_src(source: str) -> str | None:
    image_path = parse_qs(urlparse(source).query).get("url", [source])[0]
    match = re.search(r"/game/champions/(\d+)\.png", image_path)
    return match.group(1) if match else None


def read_champion_catalog(page) -> dict[str, str]:
    images = page.locator('img[src*="champions"]').evaluate_all(
        "images => images.map(image => ({"
        "alt: image.alt, "
        "source: image.getAttribute('src'), "
        "name: (image.parentElement?.parentElement?.innerText || '').trim()"
        "}))"
    )
    champions = {}
    for image in images:
        champion_id = champion_id_from_image_src(image["source"] or "")
        name = image["name"]
        if champion_id and image["alt"] not in {"ban", "pick"} and name:
            champions[champion_id] = name
    return champions


def read_champions(page, selector: str, catalog: dict[str, str]) -> list[str]:
    entries = page.locator(selector).evaluate_all(
        "groups => groups.map(group => ({"
        "text: (group.innerText || '').trim(), "
        "source: group.querySelector('img')?.getAttribute('src') || ''"
        "}))"
    )
    champions = []
    for entry in entries:
        name = entry["text"]
        if not name:
            champion_id = champion_id_from_image_src(entry["source"])
            name = catalog.get(champion_id or "", "")
        champions.append(name)
    return champions


def read_draft(page, url: str) -> dict[str, object]:
    validate_url(url)
    try:
        response = page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        page.locator(".group.blue-pick").first.wait_for(state="attached", timeout=45_000)
        page.wait_for_function(
            """() => {
                const groups = [
                    ['.group.blue-pick', false],
                    ['.group.red-pick', false],
                    ['.group.blue-ban', true],
                    ['.group.red-ban', true],
                ];
                return groups.every(([selector, imageRequired]) => {
                    const entries = Array.from(document.querySelectorAll(selector));
                    return entries.length === 5 && entries.every(entry => {
                        if (imageRequired) {
                            return Boolean(entry.querySelector('img')?.getAttribute('src'));
                        }
                        return Boolean((entry.innerText || '').trim());
                    });
                });
            }""",
            timeout=45_000,
        )
    except PlaywrightTimeoutError as error:
        raise RuntimeError(f"Les picks ou bans du draft ne sont pas complets : {url}") from error
    except PlaywrightError as error:
        raise RuntimeError(f"Impossible d'ouvrir {url} dans Edge : {error}") from error

    if response is not None and response.status >= 400:
        raise RuntimeError(f"Drafter.lol a répondu HTTP {response.status} pour {url}")

    champion_catalog = read_champion_catalog(page)
    draft = {
        "blue_team": read_team_name(page, "blue"),
        "red_team": read_team_name(page, "red"),
        "blue_picks": read_champions(page, ".group.blue-pick", champion_catalog),
        "red_picks": read_champions(page, ".group.red-pick", champion_catalog),
        "blue_bans": read_champions(page, ".group.blue-ban", champion_catalog),
        "red_bans": read_champions(page, ".group.red-ban", champion_catalog),
        "game": parse_qs(urlparse(url).query).get("game", [""])[0],
        "source_url": url,
    }

    for key in ("blue_picks", "red_picks", "blue_bans", "red_bans"):
        if len(draft[key]) != 5 or any(not champion for champion in draft[key]):
            raise RuntimeError(f"Les cinq champions n'ont pas pu être extraits pour {key} dans {url}")

    return draft


def draft_rows(draft: dict[str, object]) -> list[tuple[str, str, str, str]]:
    blue_picks = draft["blue_picks"]
    red_picks = draft["red_picks"]
    blue_bans = draft["blue_bans"]
    red_bans = draft["red_bans"]
    rows = []

    for index in range(3):
        rows.append((blue_bans[index], "BANS", "BANS", red_bans[index]))
    for index in range(3):
        number = index + 1
        rows.append((blue_picks[index], f"B{number}", f"R{number}", red_picks[index]))
    for index in range(3, 5):
        rows.append((blue_bans[index], "BANS", "BANS", red_bans[index]))
    for index in range(3, 5):
        number = index + 1
        rows.append((blue_picks[index], f"B{number}", f"R{number}", red_picks[index]))

    return rows


def add_draft_section(
    sheet,
    draft: dict[str, object],
    number: int,
    start_row: int,
) -> int:
    source_url = str(draft["source_url"])
    draft_id = urlparse(source_url).path.rstrip("/").rsplit("/", 1)[-1]
    game = str(draft["game"])
    title = f"DRAFT {number}"
    if game:
        title += f" | GAME {game}"
    if draft_id:
        title += f" | {draft_id}"

    sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=4)
    title_cell = sheet.cell(row=start_row, column=1, value=title)
    title_cell.fill = PatternFill("solid", fgColor="20262B")
    title_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    header_row = start_row + 1
    headers = ["BLUE", draft["blue_team"], draft["red_team"], "RED"]
    for column, value in enumerate(headers, start=1):
        sheet.cell(row=header_row, column=column, value=value)

    rows = draft_rows(draft)
    data_start_row = header_row + 1
    for row_index, values in enumerate(rows, start=data_start_row):
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row_index, column=column, value=value)

    header_fills = ["1F4E78", "FFD966", "A9D18E", "C00000"]
    thin_black = Side(style="thin", color="000000")
    for column in range(1, 5):
        cell = sheet.cell(row=header_row, column=column)
        cell.fill = PatternFill("solid", fgColor=header_fills[column - 1])
        cell.font = Font(
            name="Calibri",
            size=11,
            bold=True,
            color="FFFFFF" if column in (1, 4) else "000000",
        )
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_index in range(data_start_row, data_start_row + len(rows)):
        for column in range(1, 5):
            cell = sheet.cell(row=row_index, column=column)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=thin_black, right=thin_black, top=thin_black, bottom=thin_black)
            if column in (2, 3):
                cell.fill = PatternFill("solid", fgColor="000000")
                cell.font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)

    widths = (22, max(18, min(30, len(str(headers[1])) + 3)), max(18, min(30, len(str(headers[2])) + 3)), 18)
    for column, width in enumerate(widths, start=1):
        letter = get_column_letter(column)
        current_width = sheet.column_dimensions[letter].width or 0
        sheet.column_dimensions[letter].width = max(current_width, width)

    sheet.row_dimensions[start_row].height = 24
    sheet.row_dimensions[header_row].height = 24
    for row_index in range(data_start_row, data_start_row + len(rows)):
        sheet.row_dimensions[row_index].height = 22

    return data_start_row + len(rows) + 1


def create_workbook(urls: list[str], output: Path) -> None:
    if not urls or len(urls) > MAX_DRAFTS:
        raise ValueError(f"Fournis entre 1 et {MAX_DRAFTS} liens Drafter.lol.")

    with sync_playwright() as playwright:
        browser_channel = os.getenv("DRAFTER_BROWSER_CHANNEL", "msedge").strip()
        browser_name = "Chromium" if browser_channel.casefold() == "chromium" else "Microsoft Edge"
        launch_options = {
            "headless": False,
            "args": ["--disable-blink-features=AutomationControlled"],
        }
        try:
            if browser_channel.casefold() == "chromium":
                browser = playwright.chromium.launch(**launch_options)
            else:
                browser = playwright.chromium.launch(channel=browser_channel, **launch_options)
        except PlaywrightError as error:
            raise RuntimeError(f"{browser_name} est requis pour lire les drafts.") from error

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Drafts"
        next_row = 1
        try:
            for number, url in enumerate(urls, start=1):
                page = browser.new_page(viewport={"width": 1440, "height": 1100})
                page.add_init_script(
                    "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
                )
                try:
                    draft = read_draft(page, url)
                    next_row = add_draft_section(sheet, draft, number, next_row)
                finally:
                    page.close()
        finally:
            browser.close()

    sheet.freeze_panes = "A3"
    sheet.sheet_view.showGridLines = False
    sheet.print_area = f"A1:D{next_row - 2}"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extrait les picks et bans de liens Drafter.lol vers un classeur Excel."
    )
    parser.add_argument("urls", nargs="+", help="Un ou plusieurs liens /draft/ de Drafter.lol")
    parser.add_argument("-o", "--output", default="drafter_drafts.xlsx", help="Fichier Excel de sortie")
    args = parser.parse_args()

    try:
        create_workbook(args.urls, Path(args.output))
    except (ValueError, RuntimeError) as error:
        print(f"Erreur : {error}", file=sys.stderr)
        return 1

    print(f"Classeur créé : {args.output} ({len(args.urls)} draft(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())