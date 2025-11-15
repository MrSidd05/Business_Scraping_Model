import os
import re
import glob
import sys
import time
from datetime import datetime
from urllib.parse import quote_plus
from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

# ----------------------------
# Directories (duplicate_data next to extracted_data)
# ----------------------------
OUTPUT_DIR = "extracted_data"
DUP_DIR = "duplicate_data"   # placed next to OUTPUT_DIR, not inside it

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(DUP_DIR, exist_ok=True)

# ----------------------------
# Phone extraction helper
# ----------------------------
def validate_phone(raw_text):
    sequences = re.findall(r'\d+', (raw_text or ""))
    for seq in sequences:
        digits = "".join(ch for ch in seq if ch.isdigit())
        if 10 <= len(digits) <= 13:
            return digits
    return "NA"

# ----------------------------
# Duplicate-file helpers (use DUP_DIR)
# ----------------------------
def base_dup_path():
    return os.path.join(DUP_DIR, "duplicated.xlsx")

def find_latest_timestamped_dup():
    pattern = os.path.join(DUP_DIR, "duplicated_*.xlsx")
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]

def is_workbook_empty(path):
    try:
        wb = load_workbook(path)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        wb.close()
        if not rows:
            return True
        for r in rows:
            if any(cell is not None and str(cell).strip() != "" for cell in r):
                return False
        return True
    except Exception:
        return False

def read_entries_from_dup(path):
    entries = set()
    try:
        wb = load_workbook(path)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            shop = (row[1] or "").strip().lower()
            location = (row[3] or "").strip()
            entries.add((shop, location))
        wb.close()
    except Exception as e:
        print(f"Warning: couldn't read duplicates file {path}: {e}")
    return entries

def save_timestamped_dup(rows):
    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    path = os.path.join(DUP_DIR, f"duplicated_{timestamp}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(["date", "shop_name", "phone_number", "area_location", "google_maps_of_the_area"])
    for r in rows:
        ws.append(r)
    wb.save(path)
    wb.close()
    return path

def append_to_and_update_timestamp(existing_path, new_rows):
    existing_set = read_entries_from_dup(existing_path)

    filtered_rows = []
    for r in new_rows:
        shop_key = (r[1] or "").strip().lower()
        loc_key = (r[3] or "").strip()
        if (shop_key, loc_key) not in existing_set:
            filtered_rows.append(r)

    try:
        wb_existing = load_workbook(existing_path)
        ws_existing = wb_existing.active
    except Exception as e:
        print(f"Warning: couldn't open existing duplicated file: {e}. Creating fresh file.")
        return save_timestamped_dup(new_rows), len(new_rows)

    appended_count = 0
    for fr in filtered_rows:
        ws_existing.append(fr)
        appended_count += 1

    new_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    new_path = os.path.join(DUP_DIR, f"duplicated_{new_timestamp}.xlsx")
    wb_existing.save(new_path)
    wb_existing.close()

    try:
        os.remove(existing_path)
    except Exception as e:
        print(f"Warning: couldn't delete old duplicate file {existing_path}: {e}")

    return new_path, appended_count

# ----------------------------
# Load historical main entries (simple: by normalized shop name)
# ----------------------------
def load_all_previous_entries():
    """
    For compatibility and simplicity we use shop name-based historical dedupe.
    This reads any existing main_*.xlsx files and collects normalized shop names.
    """
    entries = set()
    for file in os.listdir(OUTPUT_DIR):
        if file.startswith("main_") and file.endswith(".xlsx"):
            path = os.path.join(OUTPUT_DIR, file)
            try:
                wb = load_workbook(path)
                ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    shop = (row[1] or "").strip().lower()
                    if shop:
                        entries.add(shop)
                wb.close()
            except Exception as e:
                print(f"Warning: couldn't read {path}: {e}")
    return entries

# ----------------------------
# Area validation helpers  (unchanged)
# ----------------------------
def is_obviously_invalid_area(area_input):
    if not area_input or not area_input.strip():
        return True

    s = area_input.strip()
    if s.isdigit():
        return True

    letters = re.findall(r'[A-Za-z]', s)
    if not letters:
        return True

    if len("".join(letters)) < 2:
        return True

    return False

def check_area_on_maps(area_input, headless=True, timeout_ms=8000):
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=headless)
            page = browser.new_page()
            q = quote_plus(area_input)
            search_url = f"https://www.google.com/maps/search/{q}"
            try:
                page.goto(search_url)
            except Exception:
                page.goto("https://www.google.com/maps")
                try:
                    page.fill("input#searchboxinput", area_input)
                except Exception:
                    page.evaluate("document.querySelector('input#searchboxinput').value = arguments[0];", area_input)
                page.keyboard.press("Enter")
            page.wait_for_timeout(timeout_ms)

            try:
                cards = page.locator("div.Nv2PK")
                if cards.count() and cards.count() > 0:
                    browser.close()
                    return True
            except Exception:
                pass

            try:
                content = page.content().lower()
                if "no results found" in content or "did not match" in content:
                    browser.close()
                    return False
            except Exception:
                pass

            url = page.url.lower()
            if "/search/" in url or "/place/" in url:
                browser.close()
                return True

            browser.close()
    except Exception as e:
        print(f"Warning: area validation step error: {e}")
        return False

    return False

def get_valid_area_from_user(max_attempts=2):
    attempt = 0
    while attempt < max_attempts:
        area_input = input("Enter area (spelling can be wrong, script will correct): ").strip()
        attempt += 1

        if is_obviously_invalid_area(area_input):
            print("Please enter a valid place name. Try again.")
            if attempt >= max_attempts:
                print("Unable to locate the area on Google Maps. Exiting.")
                sys.exit(1)
            continue

        if check_area_on_maps(area_input, headless=True):
            return area_input
        else:
            if attempt < max_attempts:
                print(f"The place '{area_input}' does not appear to exist in Bangalore.")
                print("Enter the correct area name within Bangalore:")
                continue
            print("Unable to locate the area. Exiting.")
            sys.exit(1)

    print("Area validation failed. Exiting.")
    sys.exit(1)

# ----------------------------
# Helper: get textual address/location for area_location column
# ----------------------------
def extract_shop_address(page):
    candidates = [
        'button[data-item-id^="address:"]',
        'button[data-item-id="address"]',
        'button[aria-label^="Address"]',
        'div.Yr7JMd-pane-hSRGPd',
        'div.IiD88e',
        'div[data-section-id="ad"]',
    ]
    for sel in candidates:
        try:
            locator = page.locator(sel).first
            if locator and locator.inner_text().strip():
                return locator.inner_text().strip()
        except Exception:
            pass
    # fallback: try to read text snippet from info pane
    try:
        content = page.locator('div.section-hero-header-title').all_text_contents()
        if content:
            return " ".join([c.strip() for c in content if c.strip()])
    except Exception:
        pass
    # fallback to current page url if nothing else
    try:
        return page.url.strip()
    except:
        return "NA"

# ----------------------------
# Helper: click Share and extract share link from dialog text
# ----------------------------
def extract_share_link_from_dialog(page):
    """
    Attempts to click the share button and extract the link text inside the dialog.
    Returns share link string or 'NA' on failure (but falls back to page.url when used).
    """
    try:
        # Try several plausible button selectors
        share_selectors = [
            'button[aria-label^="Share"]',
            'button[aria-label="Share"]',
            'button[jsaction^="pane.share"]',
            'button[aria-label*="share"]',
            'div[aria-label="Share"] button',
        ]
        clicked = False
        for sel in share_selectors:
            try:
                btn = page.locator(sel).first
                if btn and btn.is_visible():
                    btn.click()
                    clicked = True
                    break
            except Exception:
                continue

        if not clicked:
            return "NA"

        # wait for dialog
        page.wait_for_timeout(1200)
        # locate any dialog and grab text
        try:
            dialog = page.locator('div[role="dialog"]').first
            dlg_text = dialog.inner_text()
        except Exception:
            # fallback to full page
            dlg_text = page.content()

        # try to find a http(s) link in the dialog content (short maps link or full)
        m = re.search(r'(https?://\S+)', dlg_text)
        if m:
            link = m.group(1).strip().rstrip(')"\'')  # trim trailing punctuation
            # close dialog by pressing escape
            try:
                page.keyboard.press("Escape")
            except:
                pass
            return link

        # fallback: attempt to read an input's value inside the dialog
        try:
            inp = dialog.locator('input').first
            if inp:
                try:
                    val = inp.get_attribute('value') or inp.input_value()
                    if val and val.startswith('http'):
                        try:
                            page.keyboard.press("Escape")
                        except:
                            pass
                        return val
                except Exception:
                    pass
        except Exception:
            pass

        # close dialog
        try:
            page.keyboard.press("Escape")
        except:
            pass

    except Exception:
        pass

    return "NA"

# ----------------------------
# Main scraper (area_location now textual address; new column google_maps_of_the_area)
# ----------------------------
def scrape_hot_chips(area, count_needed):
    # historical dedupe by normalized shop name (simple & robust)
    historical_names = load_all_previous_entries()

    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    MAIN_FILE = os.path.join(OUTPUT_DIR, f"main_{timestamp}.xlsx")
    BASE_DUP = base_dup_path()

    # Create main file with new header (added google_maps_of_the_area)
    wb_main = Workbook()
    ws_main = wb_main.active
    ws_main.append(["date", "shop_name", "phone_number", "area_location", "google_maps_of_the_area"])
    wb_main.save(MAIN_FILE)
    wb_main.close()

    dup_entries = []
    dup_seen = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # Area-specific search (force area in query)
        query = quote_plus(f"hot chips in {area}, Bangalore")
        search_url = f"https://www.google.com/maps/search/{query}"
        try:
            page.goto(search_url)
        except Exception:
            page.goto("https://www.google.com/maps")
            try:
                page.fill("input#searchboxinput", f"hot chips in {area}, Bangalore")
            except Exception:
                page.evaluate("document.querySelector('input#searchboxinput').value = arguments[0];",
                              f"hot chips in {area}, Bangalore")
            page.keyboard.press("Enter")

        page.wait_for_timeout(5000)

        saved = 0  # count of saved rows
        seen_this_run = set()  # keep by normalized name to avoid duplicates in same run

        # retry counters when results are slow/empty
        empty_retries = 0
        empty_retries_max = 6

        while saved < count_needed:
            shops = page.locator("div.Nv2PK").all()

            # if empty results, try some recovery attempts
            if not shops:
                empty_retries += 1
                if empty_retries <= empty_retries_max:
                    try:
                        page.evaluate("window.scrollBy(0, 800);")
                    except:
                        pass
                    page.wait_for_timeout(2500)
                    shops = page.locator("div.Nv2PK").all()
                    if not shops:
                        try:
                            page.reload()
                        except:
                            pass
                        page.wait_for_timeout(3500)
                        shops = page.locator("div.Nv2PK").all()
                else:
                    print("No shop cards found after retries — stopping.")
                    break
            else:
                empty_retries = 0

            for shop in shops:
                if saved >= count_needed:
                    break
                try:
                    # name extract
                    try:
                        name = shop.locator("div.qBF1Pd").inner_text().strip()
                    except:
                        txt = shop.inner_text().strip()
                        name = txt.splitlines()[0].strip() if txt else "N/A"
                    norm_name = (name or "N/A").strip().lower()

                    # skip duplicates within same run
                    if norm_name in seen_this_run:
                        continue

                    # click card
                    try:
                        shop.click()
                    except:
                        try:
                            page.evaluate("arguments[0].click();", shop)
                        except:
                            pass

                    page.wait_for_timeout(2500)

                    # get textual address for area_location
                    try:
                        address_text = extract_shop_address(page) or ""
                    except:
                        address_text = ""

                    # extract share link (google_maps_of_the_area)
                    share_link = extract_share_link_from_dialog(page)
                    if not share_link or share_link == "NA":
                        # fallback to current page url if share link not found
                        try:
                            share_link = page.url.strip()
                        except:
                            share_link = "NA"

                    # phone extraction (unchanged)
                    phone = "NA"
                    try:
                        tel_links = page.locator('a[href^="tel:"]').all()
                        if tel_links:
                            href = tel_links[0].get_attribute("href") or ""
                            phone = validate_phone(href)
                    except:
                        pass

                    if phone == "NA":
                        try:
                            phone_btn = page.locator('button[data-item-id^="phone:"]').first
                            if phone_btn.is_visible():
                                raw_phone = phone_btn.inner_text().strip()
                                phone = validate_phone(raw_phone)
                        except:
                            pass

                    today = datetime.now().strftime("%Y-%m-%d")

                    # historical dedupe by name only (consistent with load_all_previous_entries)
                    if norm_name in historical_names:
                        dup_row = (name, phone, address_text, share_link)
                        if dup_row not in dup_seen:
                            dup_entries.append([today, name, phone, address_text, share_link])
                            dup_seen.add(dup_row)
                            print(f"[DUPLICATE] {name}")
                        seen_this_run.add(norm_name)
                        continue

                    # Save: add row with address_text in area_location and share_link in google_maps_of_the_area
                    try:
                        wb = load_workbook(MAIN_FILE)
                        ws = wb.active
                    except Exception:
                        wb = Workbook()
                        ws = wb.active
                        ws.append(["date", "shop_name", "phone_number", "area_location", "google_maps_of_the_area"])

                    ws.append([today, name, phone, address_text, share_link])
                    wb.save(MAIN_FILE)
                    wb.close()

                    saved += 1
                    seen_this_run.add(norm_name)
                    historical_names.add(norm_name)
                    print(f"Saved: {name} | Phone: {phone} | saved={saved}/{count_needed}")

                    if saved >= count_needed:
                        break

                except Exception as e:
                    print("Error while processing shop:", e)

            if saved >= count_needed:
                break

            # next page or try to load more results
            try:
                next_btn = page.locator("button[aria-label='Next']").first
                if next_btn and next_btn.is_visible():
                    next_btn.click()
                    page.wait_for_timeout(4000)
                    continue
                else:
                    # try scroll a few times
                    sc_tries = 0
                    sc_max = 4
                    loaded_more = False
                    old_count = len(shops)
                    while sc_tries < sc_max and saved < count_needed:
                        try:
                            page.evaluate("window.scrollBy(0, 900);")
                        except:
                            pass
                        page.wait_for_timeout(3000)
                        new_shops = page.locator("div.Nv2PK").all()
                        if new_shops and len(new_shops) > old_count:
                            loaded_more = True
                            break
                        sc_tries += 1
                    if loaded_more:
                        continue
                    # nothing more to load
                    print("No more result pages/cards available.")
                    break
            except Exception:
                print("Couldn't navigate to next page or load more results.")
                break

        browser.close()

    # ----------------------------
    # Duplicate handling: same behavior, but now rows include the new column
    # ----------------------------
    latest_ts_dup = find_latest_timestamped_dup()
    base_exists = os.path.exists(BASE_DUP)

    if not dup_entries:
        if not base_exists and latest_ts_dup is None:
            wb_dup = Workbook()
            ws_dup = wb_dup.active
            ws_dup.append(["date", "shop_name", "phone_number", "area_location", "google_maps_of_the_area"])
            wb_dup.save(BASE_DUP)
            wb_dup.close()
            print(f"\n🆕 Created empty base duplicate file: {BASE_DUP}")
        else:
            print("\n⭕ No duplicates found in this run. No changes to duplicate files.")
    else:
        if latest_ts_dup:
            new_path, appended = append_to_and_update_timestamp(latest_ts_dup, dup_entries)
            if base_exists and is_workbook_empty(BASE_DUP):
                try:
                    os.remove(BASE_DUP)
                except:
                    pass
            if appended > 0:
                print(f"\n⚠️ Appended {appended} new duplicate rows and updated file: {new_path}")
            else:
                print(f"\nℹ️ No new duplicate rows (already existed). Updated file: {new_path}")
        else:
            new_path = save_timestamped_dup(dup_entries)
            if base_exists and is_workbook_empty(BASE_DUP):
                try:
                    os.remove(BASE_DUP)
                except:
                    pass
            print(f"\n⚠️ Duplicate File Created: {new_path}")

    print(f"\n🎉 Main File Created: {MAIN_FILE} (rows saved this run: {saved})")

# ----------------------------
# Entrypoint
# ----------------------------
if __name__ == "__main__":
    area = get_valid_area_from_user(max_attempts=2)

    while True:
        count_input = input("How many shops to scrape?: ").strip()
        if count_input.isdigit():
            count = int(count_input)
            break
        print("❌ Please enter a valid number (no letters or special characters). Try again.")

    scrape_hot_chips(area, count)
    print("\n✅ COMPLETED SUCCESSFULLY")
