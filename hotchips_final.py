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
# Phone extraction helper
# ----------------------------
def validate_phone(raw_text):
    sequences = re.findall(r'\d+', (raw_text or ""))
    for seq in sequences:
        digits = "".join(ch for ch in seq if ch.isdigit())
        # Accept 10-13 digit numbers (India + optional country codes)
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
    """
    Return a set of (normalized_shop, location_link) present in a duplicated file.
    """
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
    """
    Append only genuinely new_rows (filtering by what's already in existing_path),
    then save as a new timestamped file and delete the old file.
    Returns (new_path, appended_count).
    """
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
        # If we can't open existing, create a new timestamped file
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

    # Remove the old existing file to avoid clutter
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
# Area validation helpers (fixed behavior)
# ----------------------------
def is_obviously_invalid_area(area_input):
    """
    Reject empty input, pure numeric input, inputs that are only punctuation,
    or super short non-meaningful inputs.
    """
    if not area_input or not area_input.strip():
        return True

    s = area_input.strip()
    # Reject if purely numeric (e.g., "10")
    if s.isdigit():
        return True

    # Reject if contains only punctuation or non-letter characters
    letters = re.findall(r'[A-Za-z]', s)
    if not letters:
        return True

    # Reject if length (letters-only) is too small (e.g., single letter)
    if len("".join(letters)) < 2:
        return True

    return False

def check_area_on_maps(area_input, headless=True, timeout_ms=8000):
    """
    Try to validate the area on Google Maps.
    Heuristics:
      - Open google maps, search for area_input and check if any result cards (div.Nv2PK) appear.
      - As a fallback check page content for common 'no results' phrases.
    Returns True if area seems valid (search results found), False otherwise.
    """
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=headless)
            page = browser.new_page()
            page.goto("https://www.google.com/maps")
            # input and search
            try:
                page.fill("input#searchboxinput", area_input)
            except Exception:
                page.evaluate("document.querySelector('input#searchboxinput').value = arguments[0];", area_input)
            page.keyboard.press("Enter")
            page.wait_for_timeout(timeout_ms)  # wait for results to load

            # Primary check: are there search result cards?
            try:
                cards = page.locator("div.Nv2PK")
                if cards.count() and cards.count() > 0:
                    browser.close()
                    return True
            except Exception:
                pass

            # Secondary check: look for 'No results' or similar text in page content
            try:
                content = page.content().lower()
                if "no results found" in content or "did not match any results" in content or "didn't match any results" in content:
                    browser.close()
                    return False
            except Exception:
                pass

            # Tertiary check: if URL looks like search or place, consider it valid
            try:
                url = page.url.lower()
                if "/search/" in url or "/place/" in url or "maps/place" in url:
                    browser.close()
                    return True
            except Exception:
                pass

            browser.close()
    except Exception as e:
        # If Playwright fails for validation, be conservative and treat as invalid
        print(f"Warning: area validation step encountered an error: {e}")
        return False

    return False

def get_valid_area_from_user(max_attempts=2):
    """
    Prompt the user for area input, validate it (local checks + google maps),
    allow up to max_attempts attempts. If still invalid, exit the program.
    Behavior:
      - If input is numeric or otherwise obviously invalid -> prompt again (no Maps check).
      - If input passes local checks but Maps says not found -> prompt again with message
        'enter the correct area name within Bangalore'. Second failure -> exit.
    """
    attempt = 0
    while attempt < max_attempts:
        prompt_msg = "Enter area (spelling can be wrong, script will correct): "
        area_input = input(prompt_msg).strip()
        attempt += 1

        # LOCAL QUICK CHECKS (reject pure numbers etc)
        if is_obviously_invalid_area(area_input):
            print("Please enter a valid place name (not a number or empty). Try again.")
            # If attempts exhausted, exit with message
            if attempt >= max_attempts:
                print("Sorry — the place you are looking for could not be located in Google Maps. Exiting.")
                sys.exit(1)
            # else loop back to prompt
            continue

        # Now check Maps — if Maps fails, ask user to enter correct area name within Bangalore
        valid_on_maps = check_area_on_maps(area_input, headless=True)
        if valid_on_maps:
            return area_input
        else:
            # If first attempt failed on maps, tell user specifically to enter area within Bangalore then retry
            if attempt < max_attempts:
                print(f"The place '{area_input}' does not appear to exist in Bangalore on Google Maps.")
                print("Enter the correct area name within Bangalore:")
                continue
            else:
                # second failure -> exit
                print("Sorry — the place you are looking for is unable to be located in Google Maps. Exiting.")
                sys.exit(1)

    # Safety exit (should not reach here)
    print("Sorry — validation failed. Exiting.")
    sys.exit(1)

# ----------------------------
# Main scraper
# ----------------------------
def scrape_hot_chips(area, count_needed):
    historical_set = load_all_previous_entries()

    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    MAIN_FILE = os.path.join(OUTPUT_DIR, f"main_{timestamp}.xlsx")
    BASE_DUP = base_dup_path()

    # create main file
    wb_main = Workbook()
    ws_main = wb_main.active
    ws_main.append(["date", "shop_name", "phone_number", "area_location"])
    wb_main.save(MAIN_FILE)
    wb_main.close()

    dup_entries = []

    # Start Playwright scraping (visible; you can change headless=False/True as desired)
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
                    # get name safely
                    try:
                        name = shop.locator("div.qBF1Pd").inner_text().strip()
                    except:
                        txt = shop.inner_text().strip()
                        name = txt.splitlines()[0].strip() if txt else "N/A"

                    # click item (safe)
                    try:
                        shop.click()
                    except:
                        page.evaluate("arguments[0].click();", shop)

                    page.wait_for_timeout(3000)

                    # phone extraction
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
                        # append to main file
                        try:
                            wb = load_workbook(MAIN_FILE)
                            ws = wb.active
                            ws.append([today, name, phone, location_link])
                            wb.save(MAIN_FILE)
                            wb.close()
                        except Exception as e:
                            print(f"Error writing to main file: {e}")

                        historical_set.add((normalized_shop, location_link))

                    num_scraped += 1
                    print(f"Saved: {name} | Phone: {phone}")

                except Exception as e:
                    print("Error while processing shop:", e)

            # pagination
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

    # ----------------------------
    # Duplicate file logic (folder-aware)
    # ----------------------------
    latest_ts_dup = find_latest_timestamped_dup()
    base_exists = os.path.exists(BASE_DUP)

    if not dup_entries:
        # No duplicates this run
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
        # Duplicates found this run
        if latest_ts_dup:
            new_path, appended = append_to_and_update_timestamp(latest_ts_dup, dup_entries)
            # delete empty base duplicated.xlsx if present
            if base_exists and is_workbook_empty(BASE_DUP):
                try:
                    os.remove(BASE_DUP)
                except Exception:
                    pass
            if appended > 0:
                print(f"\n⚠️ Appended {appended} new duplicate rows to existing file and updated timestamp: {new_path}")
            else:
                print(f"\nℹ️ No new duplicate rows to add (all duplicates already existed). Updated timestamped file: {new_path}")
        else:
            new_path = save_timestamped_dup(dup_entries)
            if base_exists and is_workbook_empty(BASE_DUP):
                try:
                    os.remove(BASE_DUP)
                except Exception:
                    pass
            print(f"\n⚠️ Duplicate File Created: {new_path}")

    print(f"\n🎉 Main File Created: {MAIN_FILE}")

# ----------------------------
# Entrypoint
# ----------------------------
if __name__ == "__main__":
    # Get a validated area from user (2 attempts). Exits if invalid twice.
    area = get_valid_area_from_user(max_attempts=2)
    try:
        count = int(input("How many shops to scrape?: "))
    except:
        count = 50
    scrape_hot_chips(area, count)
    print("\n✅ COMPLETED SUCCESSFULLY")
