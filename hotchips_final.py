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
    ws.append(["date", "shop_name", "phone_number", "area_location"])
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
# Load historical main entries
# ----------------------------
def load_all_previous_entries():
    entries = set()
    for file in os.listdir(OUTPUT_DIR):
        if file.startswith("main_") and file.endswith(".xlsx"):
            path = os.path.join(OUTPUT_DIR, file)
            try:
                wb = load_workbook(path)
                ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    shop = (row[1] or "").strip().lower()
                    location = (row[3] or "").strip()
                    entries.add((shop, location))
                wb.close()
            except Exception as e:
                print(f"Warning: couldn't read {path}: {e}")
    return entries

# ----------------------------
# Area validation helpers  (original)
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
# Main scraper (flag logic, NO closed-shop filtering)
# ----------------------------
def scrape_hot_chips(area, count_needed):
    historical_set = load_all_previous_entries()

    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    MAIN_FILE = os.path.join(OUTPUT_DIR, f"main_{timestamp}.xlsx")
    BASE_DUP = base_dup_path()

    wb_main = Workbook()
    ws_main = wb_main.active
    ws_main.append(["date", "shop_name", "phone_number", "area_location"])
    wb_main.save(MAIN_FILE)
    wb_main.close()

    dup_entries = []
    dup_seen = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # Area-specific search: "hot chips in <area>, Bangalore"
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

        flag = 0  # starts at 0, increments only when a shop is saved
        seen_this_run = set()  # avoid reprocessing same shop in same run

        # retry counters when results are slow/empty
        empty_retries = 0
        empty_retries_max = 5

        # We continue until we saved requested count or we cannot load more results
        while flag < count_needed:
            shops = page.locator("div.Nv2PK").all()

            # if no shops available, try a few recovery attempts
            if not shops:
                empty_retries += 1
                if empty_retries <= empty_retries_max:
                    # try scrolling results pane and waiting
                    try:
                        page.evaluate("window.scrollBy(0, 800);")
                    except:
                        pass
                    page.wait_for_timeout(2500)
                    shops = page.locator("div.Nv2PK").all()
                    if not shops:
                        # try reload
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

            # iterate visible cards
            for shop in shops:
                if flag >= count_needed:
                    break
                try:
                    # get shop name robustly
                    try:
                        name = shop.locator("div.qBF1Pd").inner_text().strip()
                    except:
                        txt = shop.inner_text().strip()
                        name = txt.splitlines()[0].strip() if txt else "N/A"

                    normalized_shop = (name or "N/A").strip().lower()

                    # avoid reprocessing same shop in this run
                    # use (normalized name + location_url) if available after click; so temporarily skip duplicates by name first
                    if normalized_shop in seen_this_run:
                        continue

                    # click card
                    try:
                        shop.click()
                    except:
                        try:
                            page.evaluate("arguments[0].click();", shop)
                        except:
                            pass

                    page.wait_for_timeout(3000)

                    # extract phone
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

                    location_link = page.url.strip()
                    today = datetime.now().strftime("%Y-%m-%d")

                    key = (normalized_shop, location_link)

                    # historical duplicate check
                    if key in historical_set:
                        dup_row = (name, phone, location_link)
                        if dup_row not in dup_seen:
                            dup_entries.append([today, name, phone, location_link])
                            dup_seen.add(dup_row)
                            print(f"[DUPLICATE] {name}")
                        seen_this_run.add(normalized_shop)
                        continue

                    # Not a historical duplicate -> save and increment flag
                    try:
                        wb = load_workbook(MAIN_FILE)
                        ws = wb.active
                    except Exception:
                        wb = Workbook()
                        ws = wb.active
                        ws.append(["date", "shop_name", "phone_number", "area_location"])

                    ws.append([today, name, phone, location_link])
                    wb.save(MAIN_FILE)
                    wb.close()

                    # mark saved and dedupe sets
                    flag += 1
                    seen_this_run.add(normalized_shop)
                    historical_set.add(key)
                    print(f"Saved: {name} | Phone: {phone} | flag={flag}/{count_needed}")

                    if flag >= count_needed:
                        break

                except Exception as e:
                    print("Error while processing shop:", e)

            # if we reached the target, break
            if flag >= count_needed:
                break

            # try Next page; if no Next, attempt a few scroll/retry cycles; else break
            try:
                next_btn = page.locator("button[aria-label='Next']").first
                if next_btn and next_btn.is_visible():
                    next_btn.click()
                    page.wait_for_timeout(4000)
                    continue
                else:
                    # try scroll to load more
                    sc_tries = 0
                    sc_max = 4
                    loaded_more = False
                    while sc_tries < sc_max and flag < count_needed:
                        try:
                            page.evaluate("window.scrollBy(0, 900);")
                        except:
                            pass
                        page.wait_for_timeout(3000)
                        new_shops = page.locator("div.Nv2PK").all()
                        # if new_shops number increased, we loaded more
                        if new_shops and len(new_shops) > len(shops):
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
    # Duplicate handling (unchanged)
    # ----------------------------
    latest_ts_dup = find_latest_timestamped_dup()
    base_exists = os.path.exists(BASE_DUP)

    if not dup_entries:
        if not base_exists and latest_ts_dup is None:
            wb_dup = Workbook()
            ws_dup = wb_dup.active
            ws_dup.append(["date", "shop_name", "phone_number", "area_location"])
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

    print(f"\n🎉 Main File Created: {MAIN_FILE} (rows saved this run: {flag})")

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
