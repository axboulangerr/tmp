from __future__ import annotations

import argparse
import os
import re
import sys
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urljoin, urlparse

from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
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


def read_champion_catalog(page) -> dict[str, dict[str, str]]:
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
            champions[champion_id] = {
                "name": name,
                "image_url": urljoin("https://drafter.lol", f"/game/champions/{champion_id}.png"),
            }
    return champions


def read_champions(
    page,
    selector: str,
    catalog: dict[str, dict[str, str]],
) -> tuple[list[str], list[str]]:
    entries = page.locator(selector).evaluate_all(
        "groups => groups.map(group => ({"
        "text: (group.innerText || '').trim(), "
        "source: group.querySelector('img')?.getAttribute('src') || ''"
        "}))"
    )
    names = []
    champion_ids = []
    id_by_name = {champion["name"].casefold(): champion_id for champion_id, champion in catalog.items()}
    for entry in entries:
        champion_id = champion_id_from_image_src(entry["source"])
        name = entry["text"]
        if not name:
            name = catalog.get(champion_id or "", {}).get("name", "")
        if not champion_id:
            champion_id = id_by_name.get(name.casefold(), "")
        names.append(name)
        champion_ids.append(champion_id)
    return names, champion_ids


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
    blue_picks, blue_pick_ids = read_champions(page, ".group.blue-pick", champion_catalog)
    red_picks, red_pick_ids = read_champions(page, ".group.red-pick", champion_catalog)
    blue_bans, blue_ban_ids = read_champions(page, ".group.blue-ban", champion_catalog)
    red_bans, red_ban_ids = read_champions(page, ".group.red-ban", champion_catalog)
    draft = {
        "blue_team": read_team_name(page, "blue"),
        "red_team": read_team_name(page, "red"),
        "blue_picks": blue_picks,
        "red_picks": red_picks,
        "blue_bans": blue_bans,
        "red_bans": red_bans,
        "blue_pick_ids": blue_pick_ids,
        "red_pick_ids": red_pick_ids,
        "blue_ban_ids": blue_ban_ids,
        "red_ban_ids": red_ban_ids,
        "champion_image_urls": {
            champion_id: champion["image_url"]
            for champion_id, champion in champion_catalog.items()
        },
        "game": parse_qs(urlparse(url).query).get("game", [""])[0],
        "source_url": url,
    }

    for key in ("blue_picks", "red_picks", "blue_bans", "red_bans"):
        image_key = key.removesuffix("s") + "_ids"
        if (
            len(draft[key]) != 5
            or any(not champion for champion in draft[key])
            or any(champion_id not in draft["champion_image_urls"] for champion_id in draft[image_key])
        ):
            raise RuntimeError(f"Les cinq champions n'ont pas pu être extraits pour {key} dans {url}")

    return draft


def draft_rows(draft: dict[str, object]) -> list[tuple[str, str, str, str, str, str]]:
    blue_picks = draft["blue_picks"]
    red_picks = draft["red_picks"]
    blue_bans = draft["blue_bans"]
    red_bans = draft["red_bans"]
    blue_pick_ids = draft["blue_pick_ids"]
    red_pick_ids = draft["red_pick_ids"]
    blue_ban_ids = draft["blue_ban_ids"]
    red_ban_ids = draft["red_ban_ids"]
    rows = []

    for index in range(3):
        rows.append((blue_bans[index], "BANS", "BANS", red_bans[index], blue_ban_ids[index], red_ban_ids[index]))
    for index in range(3):
        number = index + 1
        rows.append((blue_picks[index], f"B{number}", f"R{number}", red_picks[index], blue_pick_ids[index], red_pick_ids[index]))
    for index in range(3, 5):
        rows.append((blue_bans[index], "BANS", "BANS", red_bans[index], blue_ban_ids[index], red_ban_ids[index]))
    for index in range(3, 5):
        number = index + 1
        rows.append((blue_picks[index], f"B{number}", f"R{number}", red_picks[index], blue_pick_ids[index], red_pick_ids[index]))

    return rows


def champion_thumbnail(page, image_url: str, cache: dict[str, bytes]) -> bytes:
    if image_url not in cache:
        response = page.context.request.get(image_url, timeout=30_000)
        if not response.ok:
            raise RuntimeError(f"Impossible de télécharger l'image du champion : HTTP {response.status}")

        with PILImage.open(BytesIO(response.body())) as image:
            thumbnail = image.convert("RGBA")
            thumbnail.thumbnail((32, 32))
            buffer = BytesIO()
            thumbnail.save(buffer, format="PNG")
            cache[image_url] = buffer.getvalue()

    return cache[image_url]


def add_draft_section(
    sheet,
    draft: dict[str, object],
    number: int,
    start_row: int,
    page,
    image_cache: dict[str, bytes],
) -> int:
    source_url = str(draft["source_url"])
    draft_id = urlparse(source_url).path.rstrip("/").rsplit("/", 1)[-1]
    game = str(draft["game"])
    title = f"DRAFT {number}"
    if game:
        title += f" | GAME {game}"
    if draft_id:
        title += f" | {draft_id}"

    sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
    title_cell = sheet.cell(row=start_row, column=1, value=title)
    title_cell.fill = PatternFill("solid", fgColor="20262B")
    title_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    header_row = start_row + 1
    headers = ["ICON", "BLUE", draft["blue_team"], draft["red_team"], "RED", "ICON"]
    for column, value in enumerate(headers, start=1):
        sheet.cell(row=header_row, column=column, value=value)

    rows = draft_rows(draft)
    data_start_row = header_row + 1
    for row_index, values in enumerate(rows, start=data_start_row):
        blue_name, blue_side, red_side, red_name, blue_id, red_id = values
        for column, value in ((2, blue_name), (3, blue_side), (4, red_side), (5, red_name)):
            sheet.cell(row=row_index, column=column, value=value)
        for column, champion_id in ((1, blue_id), (6, red_id)):
            image_url = draft["champion_image_urls"][champion_id]
            icon = ExcelImage(BytesIO(champion_thumbnail(page, image_url, image_cache)))
            icon.width = 32
            icon.height = 32
            sheet.add_image(icon, f"{get_column_letter(column)}{row_index}")

    header_fills = ["274C77", "1F4E78", "FFD966", "A9D18E", "C00000", "7F1D1D"]
    thin_black = Side(style="thin", color="000000")
    for column in range(1, 7):
        cell = sheet.cell(row=header_row, column=column)
        cell.fill = PatternFill("solid", fgColor=header_fills[column - 1])
        cell.font = Font(
            name="Calibri",
            size=11,
            bold=True,
            color="FFFFFF" if column in (1, 2, 5, 6) else "000000",
        )
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_index in range(data_start_row, data_start_row + len(rows)):
        for column in range(1, 7):
            cell = sheet.cell(row=row_index, column=column)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=thin_black, right=thin_black, top=thin_black, bottom=thin_black)
            if column in (3, 4):
                cell.fill = PatternFill("solid", fgColor="000000")
                cell.font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)

    widths = (
        6,
        22,
        max(18, min(30, len(str(headers[2])) + 3)),
        max(18, min(30, len(str(headers[3])) + 3)),
        22,
        6,
    )
    for column, width in enumerate(widths, start=1):
        letter = get_column_letter(column)
        current_width = sheet.column_dimensions[letter].width or 0
        sheet.column_dimensions[letter].width = max(current_width, width)

    sheet.row_dimensions[start_row].height = 24
    sheet.row_dimensions[header_row].height = 24
    for row_index in range(data_start_row, data_start_row + len(rows)):
        sheet.row_dimensions[row_index].height = 30

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
        image_cache: dict[str, bytes] = {}
        try:
            for number, url in enumerate(urls, start=1):
                page = browser.new_page(viewport={"width": 1440, "height": 1100})
                page.add_init_script(
                    "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
                )
                try:
                    draft = read_draft(page, url)
                    next_row = add_draft_section(
                        sheet,
                        draft,
                        number,
                        next_row,
                        page,
                        image_cache,
                    )
                finally:
                    page.close()
        finally:
            browser.close()

    sheet.freeze_panes = "A3"
    sheet.sheet_view.showGridLines = False
    sheet.print_area = f"A1:F{next_row - 2}"
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