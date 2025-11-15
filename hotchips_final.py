import os
import re
import glob
import sys
from datetime import datetime
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
# Phone extraction helper (micro-optimized: compile regex once)
# ----------------------------
_PHONE_RE = re.compile(r'\d+')

def validate_phone(raw_text):
    sequences = _PHONE_RE.findall(raw_text or "")
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
    """
    Small optimization: use max() instead of sorting the entire list.
    """
    pattern = os.path.join(DUP_DIR, "duplicated_*.xlsx")
    files = glob.glob(pattern)
    if not files:
        return None
    return max(files, key=os.path.getmtime)

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
# Area validation helpers (kept behavior exactly the same)
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
    """
    Original behavior preserved exactly: this launches Playwright, navigates, and waits.
    """
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=headless)
            page = browser.new_page()
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

def check_area_on_maps_page(page, area_input, timeout_ms=8000):
    """
    Same checks as check_area_on_maps but operates on an existing Playwright page instance.
    This allows reusing a single browser/page to avoid repeated launches in get_valid_area_from_user.
    Functionality and waits are unchanged.
    """
    try:
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
                return True
        except Exception:
            pass

        try:
            content = page.content().lower()
            if "no results found" in content or "did not match" in content:
                return False
        except Exception:
            pass

        url = page.url.lower()
        if "/search/" in url or "/place/" in url:
            return True
    except Exception as e:
        print(f"Warning: area validation (page) error: {e}")
        return False

    return False

def get_valid_area_from_user(max_attempts=2):
    """
    Micro-optimized: reuse a single Playwright browser/page across attempts.
    Behavior (waits & selectors) remains identical to original.
    """
    attempt = 0
    # We'll launch the browser once, use for up to max_attempts, then close
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        while attempt < max_attempts:
            area_input = input("Enter area (spelling can be wrong, script will correct): ").strip()
            attempt += 1

            if is_obviously_invalid_area(area_input):
                print("Please enter a valid place name. Try again.")
                if attempt >= max_attempts:
                    print("Unable to locate the area on Google Maps. Exiting.")
                    browser.close()
                    sys.exit(1)
                continue

            # use the page-based checker which has identical checks/waits as the original
            if check_area_on_maps_page(page, area_input, timeout_ms=8000):
                browser.close()
                return area_input
            else:
                if attempt < max_attempts:
                    print(f"The place '{area_input}' does not appear to exist in Bangalore.")
                    print("Enter the correct area name within Bangalore:")
                    continue
                print("Unable to locate the area. Exiting.")
                browser.close()
                sys.exit(1)

        print("Area validation failed. Exiting.")
        browser.close()
        sys.exit(1)

# ----------------------------
# Main scraper (kept unchanged in behavior)
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

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        page.goto("https://www.google.com/maps")
        page.fill("input#searchboxinput", area)
        page.keyboard.press("Enter")
        page.wait_for_timeout(5000)

        page.fill("input#searchboxinput", "")
        page.fill("input#searchboxinput", "hot chips near me")
        page.keyboard.press("Enter")
        page.wait_for_timeout(6000)

        num_scraped = 0
        while num_scraped < count_needed:
            shops = page.locator("div.Nv2PK").all()
            if not shops:
                break

            for shop in shops:
                if num_scraped >= count_needed:
                    break
                try:
                    try:
                        name = shop.locator("div.qBF1Pd").inner_text().strip()
                    except:
                        txt = shop.inner_text().strip()
                        name = txt.splitlines()[0].strip() if txt else "N/A"

                    try:
                        shop.click()
                    except:
                        page.evaluate("arguments[0].click();", shop)

                    page.wait_for_timeout(3000)

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
                    normalized_shop = (name or "N/A").strip().lower()

                    if (normalized_shop, location_link) in historical_set:
                        dup_entries.append([today, name, phone, location_link])
                        print(f"[DUPLICATE] {name}")
                    else:
                        wb = load_workbook(MAIN_FILE)
                        ws = wb.active
                        ws.append([today, name, phone, location_link])
                        wb.save(MAIN_FILE)
                        wb.close()
                        historical_set.add((normalized_shop, location_link))

                    num_scraped += 1
                    print(f"Saved: {name} | Phone: {phone}")

                except Exception as e:
                    print("Error while processing shop:", e)

            try:
                next_btn = page.locator("button[aria-label='Next']").first
                if next_btn.is_visible():
                    next_btn.click()
                    page.wait_for_timeout(4000)
                    continue
            except:
                pass

            break

        browser.close()

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

    print(f"\n🎉 Main File Created: {MAIN_FILE}")

# ----------------------------
# Entrypoint
# ----------------------------
if __name__ == "__main__":
    area = get_valid_area_from_user(max_attempts=2)

    # ----------------------------
    # UPDATED LOGIC (your request)
    # ----------------------------
    while True:
        count_input = input("How many shops to scrape?: ").strip()
        if count_input.isdigit():
            count = int(count_input)
            break
        print("❌ Please enter a valid number (no letters or special characters). Try again.")

    scrape_hot_chips(area, count)
    print("\n✅ COMPLETED SUCCESSFULLY")
