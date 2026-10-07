"""
SLAM Loco Equipment Automation
================================
Kal wale "Loco Task Monitoring" project jaisa hi pattern:
  - User khud login karta hai (SLAM par math captcha hota hai, isliye login
    automate nahi kiya gaya - safest yahi hai).
  - Script Google Sheet se data padhti hai (CSV export, koi API key nahi chahiye).
  - Har Loco ke liye equipment list (MRB/OCB/TMB x Position 1/2) kholti hai.
  - Equipment link par click karne se NAYA TAB khulta hai -> script us tab
    par switch karke form bharti hai -> Save -> tab band -> wapas list wale
    tab par -> agla equipment.

============================================================================
JARURI: Field selectors (labels) screenshots se liye gaye hain. Agar SLAM
ke actual page par label ka exact text thoda alag nikla (jaise "Eq.Number"
vs "Eq. Number"), to FIELD_LABELS dictionary neeche update kar dena.
Pehli baar script chalane par ek loco, ek equipment ke saath TEST zaroor
karein.
============================================================================
"""

import time
import sys
import re
import pandas as pd
import requests
from io import StringIO

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    InvalidElementStateException,
    UnexpectedAlertPresentException,
    NoAlertPresentException,
    NoSuchWindowException,
)
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service


# ============================================================================
# CONFIG - yahan apni details daalein
# ============================================================================

# Google Sheet ID (poora link nahi, sirf /d/ aur /edit ke beech wala part)
SHEET_ID = "1hzj67LKBl15xVTc7E43APChoSzMZfNAB"

# Agar sheet mein specific tab/gid use karna hai to yahan daalein, warna
# pehla/default sheet padhega. Sheet URL mein #gid=XXXX dekh kar pata chalega.
SHEET_GID = "1093714984"  # "Sheet1" tab - jahan 30541 ki poori details hain

# SLAM site
LOGIN_URL = "https://slam.indianrail.gov.in/SLAM/Login.aspx"
ADD_EQUIPMENT_URL = "https://slam.indianrail.gov.in/SLAM/EquipmentRegister/LocoEquipmentAdd.aspx"

LOCATION_VALUE = "RPME"   # aapke workflow mein fixed location
SELECT_FOR_VALUE = "Home"

WAIT_TIMEOUT = 20  # seconds


# ============================================================================
# STEP 1: Google Sheet se data padhna
# ============================================================================

def load_sheet_data():
    if SHEET_GID:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={SHEET_GID}"
    else:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

    resp = requests.get(url)
    resp.raise_for_status()
    df = pd.read_csv(StringIO(resp.text), dtype=str)

    # Column names ke aage/peeche ke spaces hata dena (jaise "Eq. Number (M) ")
    df.columns = [c.strip() for c in df.columns]

    # Sab values ko string bana kar strip karna, NaN ko empty string
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": ""})

    print(f"[OK] Sheet se {len(df)} rows padhi gayi.")
    print(f"[DEBUG] Sheet ke columns: {df.columns.tolist()}")
    print("[DEBUG] Sheet ka poora data:")
    print(df.to_string())
    return df


def is_blank(value):
    """Khali ya 'NA' jaisi values ko blank treat karo."""
    if value is None:
        return True
    v = str(value).strip().upper()
    return v in ("", "NA", "N/A", "-", "NONE")


def get_row_for(df, loco_no, eq_name, position):
    """Sheet mein Loco No + Eq Name + Position match karke row dhoondhna.

    Kuch sheets mein 'Eq Name' column mein Position already included hota hai
    (jaise 'MRB1', 'TMB2') aur alag 'Position' column khali hota hai. Isliye
    yahan Eq Name ko letters+digits mein split karke bhi try karte hain.
    """
    loco_no = str(loco_no).strip()
    eq_name = eq_name.upper().strip()
    position = str(position).strip()

    df_loco = df[df["Loco No"].astype(str).str.strip() == loco_no]

    for _, r in df_loco.iterrows():
        sheet_eq_raw = str(r["Eq Name"]).upper().strip()
        sheet_pos_raw = str(r.get("Position", "")).strip()

        # Case 1: Eq Name mein sirf naam hai (jaise "MRB"), Position alag column mein hai
        if sheet_eq_raw == eq_name and sheet_pos_raw == position:
            return r

        # Case 2: Eq Name mein naam+position combined hai (jaise "MRB1")
        m = re.match(r"^([A-Z]+)\s*0*(\d+)$", sheet_eq_raw)
        if m:
            combined_name = m.group(1)
            combined_pos = m.group(2)
            if combined_name == eq_name and (
                combined_pos == position.lstrip("0") or combined_pos == position
            ):
                return r

    return None


# ============================================================================
# STEP 2: Selenium helpers - label text se input/select dhoondhna
# ============================================================================
# SLAM ke ASP.NET generated element IDs random/lambe hote hain (jaise
# ctl00_ContentPlaceHolder1_txtXYZ), isliye hum label ke TEXT se dhoondhte
# hain - ye zyada reliable hai. Screenshot mein form ek table hai jisme
# left column mein label aur right column mein input/select hai.

def find_input_by_label(driver, label_text):
    # Pehle table-row pattern try karo (Add Loco Equipment ke form jaisa)
    xpath_table = (
        f"//td[normalize-space(text())='{label_text}']"
        f"/following-sibling::td[1]//input"
    )
    try:
        return WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, xpath_table))
        )
    except TimeoutException:
        pass
    # Fallback: label text ke turant baad document-order mein pehla input
    xpath_generic = f"//*[normalize-space(text())='{label_text}']/following::input[1]"
    return WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, xpath_generic))
    )


def find_select_by_label(driver, label_text):
    xpath_table = (
        f"//td[normalize-space(text())='{label_text}']"
        f"/following-sibling::td[1]//select"
    )
    try:
        return WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, xpath_table))
        )
    except TimeoutException:
        pass
    xpath_generic = f"//*[normalize-space(text())='{label_text}']/following::select[1]"
    return WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, xpath_generic))
    )


# ----------------------------------------------------------------------------
# Frame handling: SLAM (jaise kal ki site) content kabhi-kabhi <iframe> ke
# andar hota hai. Ye function default content mein dhoondhta hai, na mile to
# har iframe ke andar check karta hai aur us frame par switch kar deta hai.
# ----------------------------------------------------------------------------

def text_present(driver, text):
    try:
        driver.find_element(By.XPATH, f"//*[contains(normalize-space(text()),'{text}')]")
        return True
    except NoSuchElementException:
        return False


def switch_to_frame_containing(driver, text, max_wait=10):
    driver.switch_to.default_content()
    end_time = time.time() + max_wait
    while time.time() < end_time:
        if text_present(driver, text):
            return True  # already in the right place (default content)
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for f in frames:
            driver.switch_to.default_content()
            try:
                driver.switch_to.frame(f)
                if text_present(driver, text):
                    return True
            except Exception:
                continue
        driver.switch_to.default_content()
        time.sleep(0.5)
    driver.switch_to.default_content()
    return False


def debug_list_selects(driver):
    """Debugging ke liye: page par mile saare <select> elements ki info print karo."""
    print("    [DEBUG] Page par mile <select> elements:")
    driver.switch_to.default_content()
    _debug_list_selects_in_current_context(driver, "main page")
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for i, f in enumerate(frames):
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(f)
            _debug_list_selects_in_current_context(driver, f"iframe #{i}")
        except Exception as e:
            print(f"      (iframe #{i} access error: {e})")
    driver.switch_to.default_content()


def _debug_list_selects_in_current_context(driver, context_name):
    selects = driver.find_elements(By.TAG_NAME, "select")
    if not selects:
        print(f"      {context_name}: koi <select> nahi mila")
        return
    for s in selects:
        try:
            print(f"      {context_name}: id='{s.get_attribute('id')}' "
                  f"name='{s.get_attribute('name')}' "
                  f"options={[o.text for o in Select(s).options][:5]}")
        except Exception:
            pass


def set_text(driver, label_text, value, wait_enabled=False):
    if is_blank(value):
        return
    try:
        field = find_input_by_label(driver, label_text)
        if wait_enabled:
            # Kuch fields (jaise "Warranty Upto") disabled hote hain jab tak
            # dropdown "YES" select na ho - JS se enable hone ka thoda wait.
            try:
                WebDriverWait(driver, 5).until(lambda d: field.is_enabled())
            except TimeoutException:
                pass
        if not field.is_enabled():
            print(f"    [SKIP] '{label_text}' field abhi bhi disabled hai, skip kar rahe hain.")
            return
        try:
            field.clear()
            field.send_keys(value)
            print(f"    [text] {label_text} = {value}")
        except InvalidElementStateException:
            # Date fields (Date of Commission/Manufacture/Recieve, Warranty
            # Upto, AMC Upto) SLAM par 'readonly' hote hain - inpar sirf
            # calendar-picker se click karke value set hoti hai, typing se
            # nahi. Selenium clear()/send_keys() isliye fail hota hai. Fix:
            # JS se seedha value daal kar 'input'/'change'/'blur' events fire
            # karo, taaki SLAM ki JS validation bhi is value ko maane.
            driver.execute_script(
                "arguments[0].removeAttribute('readonly');"
                "arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));"
                "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));"
                "arguments[0].dispatchEvent(new Event('blur', {bubbles: true}));",
                field, value,
            )
            print(f"    [text-js] {label_text} = {value} (readonly field, JS se set kiya)")
    except (NoSuchElementException, TimeoutException):
        print(f"    [SKIP] Field not found for label: {label_text}")
    except Exception as e:
        print(f"    [SKIP] '{label_text}' field mein issue aaya ({type(e).__name__}): {e}")


def set_dropdown(driver, label_text, value):
    if is_blank(value):
        return
    try:
        dropdown = find_select_by_label(driver, label_text)
        sel = Select(dropdown)

        # 1) Exact match
        try:
            sel.select_by_visible_text(value)
            print(f"    [dropdown] {label_text} = {value}")
            return
        except NoSuchElementException:
            pass

        # 2) Case-insensitive match
        for opt in sel.options:
            if opt.text.strip().upper() == value.strip().upper():
                sel.select_by_visible_text(opt.text)
                print(f"    [dropdown] {label_text} = {opt.text} "
                      f"(case-insensitive match, sheet mein '{value}' tha)")
                return

        # 3) Look-alike character mismatch - jaise 'ICO' (letter O) vs SLAM
        # ke dropdown mein 'IC0' (digit zero). Aksar sheet aur site ke
        # schedule codes mein O/0 ka typo/confusion hota hai.
        target_norm = value.strip().upper().replace('O', '0')
        for opt in sel.options:
            if opt.text.strip().upper().replace('O', '0') == target_norm:
                sel.select_by_visible_text(opt.text)
                print(f"    [dropdown] {label_text} = {opt.text} "
                      f"(O/0 mismatch fix kiya, sheet mein '{value}' tha)")
                return

        # Kuch bhi match nahi hua - dropdown '--Select--' par hi reh gaya hai,
        # is form ko save karne se pehle sheet ki value check karni chahiye.
        print(f"    [WARN] Dropdown option '{value}' nahi mila '{label_text}' mein - "
              f"KUCH BHI SELECT NAHI HUA, yeh field khaali reh gayi hai! "
              f"Available options: {[o.text for o in sel.options]}")
    except (NoSuchElementException, TimeoutException):
        print(f"    [SKIP] Dropdown not found for label: {label_text}")
    except Exception as e:
        print(f"    [SKIP] '{label_text}' dropdown mein issue aaya ({type(e).__name__}): {e}")


def dismiss_alert_if_present(driver, accept=True):
    """SLAM kabhi-kabhi mandatory-field JS alert dikhata hai (jaise 'Please
    enter Warranty Upto date'). Agar isko handle na karein to Selenium ki
    agli command 'unexpected alert open' error de kar poora script crash kar
    degi. Ye function alert hoto usko accept/dismiss kar deta hai aur uska
    text return karta hai (warna None)."""
    try:
        alert = driver.switch_to.alert
        text = alert.text
        if accept:
            alert.accept()
        else:
            alert.dismiss()
        print(f"    [ALERT] Popup mila aur band kiya: {text!r}")
        return text
    except NoAlertPresentException:
        return None
    except Exception:
        return None


def click_go_button(driver, timeout=WAIT_TIMEOUT):
    """Naye tab par pehle sirf 'Available Type: New' + 'Go' dikhta hai - baaki
    saare fields (Provision Date, Schedule, Model, Eq.Number, dates, waranty,
    etc.) DOM mein tab tak exist hi nahi karte jab tak ye 'Go' button click na
    ho. Isko click karna zaroori hai, warna form bharna fail hoga."""
    try:
        go_btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='Go' or @value='GO']"))
        )
        go_btn.click()
        print("    [INFO] 'Go' button (Available Type) click kiya.")
        return True
    except (TimeoutException, NoSuchElementException) as e:
        print(f"    [WARN] 'Go' button (Available Type) nahi mila/click nahi hua: "
              f"{type(e).__name__}: {e}")
        return False


# ============================================================================
# STEP 3: Ek equipment form bharna (naye tab ke andar)
# ============================================================================

def fill_equipment_form(driver, row):
    """`row` ek pandas Series hai jisme us equipment ki sheet-row hai."""

    # Available Type by default 'New' hi hota hai, isliye chhedte nahi.

    set_text(driver, "Provision Date", row.get("Provision Date", ""))
    set_dropdown(driver, "Schedule", row.get("Schedule", ""))
    # Schedule select karne se ASP.NET postback trigger ho sakta hai (onchange
    # event se hi, "Go" button dabane ki zarurat nahi - wo button "Available
    # Type" ke liye hai jo already "New" hai). Postback settle hone ka wait:
    time.sleep(1.5)

    set_dropdown(driver, "Model", row.get("Model", ""))
    set_text(driver, "Eq.Number", row.get("Eq. Number (M)", row.get("Eq. Number", "")))
    set_text(driver, "Location Eq.Number", row.get("Location Eq. Number", row.get("Eq. Number (L)", "")))
    set_text(driver, "Date of Commission", row.get("DoC", ""))
    set_text(driver, "Date of Manufacture", row.get("DoM", ""))
    set_text(driver, "Date of Recieve", row.get("Date of Receive", ""))  # note: SLAM spelling "Recieve"

    warranty_val = row.get("Warranty", "")
    set_dropdown(driver, "Warranty", warranty_val)
    if not is_blank(warranty_val) and warranty_val.strip().upper() == "YES":
        time.sleep(0.5)
        set_text(driver, "Warranty Upto(If Avbl)", row.get("Warranty Upto", ""), wait_enabled=True)

    amc_val = row.get("AMC", "")
    set_dropdown(driver, "AMC", amc_val)
    if not is_blank(amc_val) and amc_val.strip().upper() == "YES":
        time.sleep(0.5)
        set_text(driver, "AMC Upto (If Avbl)", row.get("AMC Upto", ""), wait_enabled=True)

    set_dropdown(driver, "Eq.Make", row.get("Make", ""))
    set_text(driver, "Remarks", row.get("Remarks", ""))

    # ---- SAVE ----
    try:
        save_btn = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='Save']"))
        )
        save_btn.click()
        time.sleep(2)
        print("    [SAVE] Form save kar diya.")
    except Exception as e:
        print(f"    [ERROR] Save button click nahi hua: {type(e).__name__}: {e}")


# ============================================================================
# STEP 4: Ek Loco ke saare equipment process karna
# ============================================================================

def process_loco(driver, df, loco_no):
    print(f"\n=== Loco {loco_no} processing shuru ===")

    driver.get(ADD_EQUIPMENT_URL)
    time.sleep(2)

    # Content shayad iframe ke andar ho - "Loco" text jahan mile wahi switch karo
    found = switch_to_frame_containing(driver, "Loco")
    if not found:
        print("[WARN] 'Loco' text kahin nahi mila (na main page, na iframe mein). "
              "Page load hone mein time lag raha ho sakta hai, ya login expire ho gaya.")
        debug_list_selects(driver)
        return

    # "Select For" hamesha default "Home" hi rehta hai - chhedte nahi, sirf
    # best-effort confirm karte hain. Agar already Home hai to kuch nahi hoga.
    try:
        select_for = Select(driver.find_element(By.ID, "Content1_ddlLocoShed"))
        if select_for.first_selected_option.text.strip() != SELECT_FOR_VALUE:
            select_for.select_by_visible_text(SELECT_FOR_VALUE)
    except (NoSuchElementException, TimeoutException):
        pass  # already default Home hai, ignore karo

    time.sleep(1)

    # Loco dropdown - ye zaroori hai (known element ID: Content1_ddlLoco)
    try:
        loco_dd = Select(WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.ID, "Content1_ddlLoco"))
        ))
        loco_dd.select_by_visible_text(str(loco_no))
    except (NoSuchElementException, TimeoutException):
        print(f"[ERROR] Loco dropdown (id=Content1_ddlLoco) nahi mila. Debug info:")
        debug_list_selects(driver)
        return
    except Exception as e:
        # select_by_visible_text fail ho sakta hai agar Loco is shed ke Home list mein nahi hai
        print(f"[ERROR] Loco {loco_no} ko dropdown mein select nahi kar paya: {e}")
        try:
            options = [o.text for o in loco_dd.options]
            print(f"    Available loco options ({len(options)} total): {options}")
        except Exception:
            pass
        return

    # Go button (usi frame ke andar)
    go_btn = WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@value='Go' or @value='GO']"))
    )
    go_btn.click()
    time.sleep(2)

    original_window = driver.current_window_handle

    # Equipment list table ke links: "Filter for MRB", "Filter for OCB", "Filter for TMB"
    equipment_links = driver.find_elements(By.PARTIAL_LINK_TEXT, "Filter for")
    print(f"[INFO] {len(equipment_links)} equipment links mile is loco ke liye.")

    total_links = len(equipment_links)

    # Table ke header row se "Position" column ka exact index pata karo,
    # taaki chahe column order kuch bhi ho, sahi cell padhi jaye.
    position_col_index = 3  # default fallback (S.No, Eq.Name, Eq.Model, Eq.Position)
    if total_links > 0:
        try:
            table_el = equipment_links[0].find_element(By.XPATH, "./ancestor::table[1]")
            header_cells = table_el.find_elements(By.XPATH, ".//tr[1]/th | .//tr[1]/td")
            for i, h in enumerate(header_cells):
                if "position" in h.text.strip().lower():
                    position_col_index = i
                    break
            print(f"    [DEBUG] Position column index detect hua: {position_col_index} "
                  f"(headers: {[h.text.strip() for h in header_cells]})")
        except Exception as e:
            print(f"    [WARN] Header row detect nahi hua, default index {position_col_index} use karenge: {e}")

    # JARURI: Ek equipment save hone ke baad SLAM usko is 'pending' list se
    # hata deta hai - isliye list HAR save ke baad shrink hoti hai. Agar hum
    # fixed index (0,1,2...) se access karein to list shrink hone par index
    # shift ho jata hai aur galat/agla item skip ho jata hai (jaisa OCB1 aur
    # TMB1 ke saath hua tha). Fix: index ke bajaye, har baar fresh list mein
    # se pehla aisa link dhoondhte hain jo abhi tak 'handle' nahi hua -
    # (eq_name, position) ko track karke.
    handled_keys = set()
    max_iterations = total_links + 5  # safety cap, infinite loop se bachne ke liye

    for _ in range(max_iterations):
        equipment_links = driver.find_elements(By.PARTIAL_LINK_TEXT, "Filter for")
        if not equipment_links:
            break

        link = None
        eq_name = position = ""
        for candidate in equipment_links:
            try:
                row_el = candidate.find_element(By.XPATH, "./ancestor::tr[1]")
                cells = row_el.find_elements(By.TAG_NAME, "td")
                cand_eq_name = candidate.text.strip().replace("Filter for", "").strip()
                cand_position = cells[position_col_index].text.strip() if position_col_index < len(cells) else ""
            except (StaleElementReferenceException, NoSuchElementException):
                continue
            if (cand_eq_name, cand_position) not in handled_keys:
                link, eq_name, position = candidate, cand_eq_name, cand_position
                break

        if link is None:
            # Jitne bhi links list mein hain, sab pehle hi handle ho chuke hain
            break

        handled_keys.add((eq_name, position))
        print(f"\n  -> Equipment: {eq_name}, Position: {position}")

        sheet_row = get_row_for(df, loco_no, eq_name, position)
        if sheet_row is None:
            print("    [SKIP] Sheet mein is combination ki row nahi mili.")
            continue

        # Agar sheet row mein zaroori fields (Eq Number, Make) bhi blank/NA hain
        # to is equipment ko chhod do (kuch nahi bharna).
        if is_blank(sheet_row.get("Eq. Number (M)", "")) and is_blank(sheet_row.get("Eq. Number", "")):
            print("    [SKIP] Eq. Number blank/NA hai sheet mein, form nahi bharenge.")
            continue

        # Naya tab click se khulta hai
        before_windows = driver.window_handles
        link.click()
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            lambda d: len(d.window_handles) > len(before_windows)
        )
        new_window = [w for w in driver.window_handles if w not in before_windows][0]
        driver.switch_to.window(new_window)
        time.sleep(1.5)

        # Pehle "Available Type" wala Go button click karo - tabhi form ke
        # baaki fields load honge. Kabhi-kabhi (jaise OCB equipment mein
        # dekha gaya) pehli koshish mein postback slow ya miss ho jata hai,
        # isliye zaroorat pade to ek dobara try bhi karte hain.
        form_ready = False
        for attempt in (1, 2):
            switch_to_frame_containing(driver, "Available Type")
            go_clicked = click_go_button(driver)
            if go_clicked:
                time.sleep(2)  # postback settle hone ka wait
            form_ready = switch_to_frame_containing(driver, "Provision Date")
            if form_ready:
                break
            print(f"    [WARN] Attempt {attempt}: form fields load nahi hue"
                  + (", dobara try kar rahe hain..." if attempt == 1 else "."))
            time.sleep(1.5)

        if not form_ready:
            print("    [ERROR] 'Go' click ke baad bhi form fields load nahi hue - "
                  "is equipment ko skip kar rahe hain.")
        else:
            try:
                fill_equipment_form(driver, sheet_row)
            except UnexpectedAlertPresentException:
                dismiss_alert_if_present(driver)
                print("    [ERROR] Form bharte waqt ek popup/alert aa gaya, "
                      "is equipment ko skip kar diya.")
            except Exception as e:
                print(f"    [ERROR] Form bharte waqt issue: {e}")

        # Agar koi mandatory-field alert khula reh gaya ho (jaise 'Please
        # enter Warranty Upto date'), pehle usko band karo - warna tab close
        # karna bhi fail ho kar poora loco crash kar dega.
        dismiss_alert_if_present(driver)

        # Agar SLAM khud "Saved Successfully" ke baad tab/popup band kar deta
        # hai, to hamara apna driver.close() 'no such window' error dega -
        # isliye pehle check karte hain ki wo window abhi bhi khuli hai.
        try:
            if new_window in driver.window_handles:
                driver.close()
        except (UnexpectedAlertPresentException, NoSuchWindowException):
            dismiss_alert_if_present(driver)

        # Wapas original (list wale) tab par switch karo. Agar wo bhi kisi
        # wajah se nahi mila (rare case), to jo bhi window bachi hai usi par
        # switch kar lo taaki loop crash na ho.
        try:
            driver.switch_to.window(original_window)
        except NoSuchWindowException:
            remaining = driver.window_handles
            if remaining:
                print("    [WARN] Original list-tab nahi mila, available "
                      "window par switch kar rahe hain.")
                driver.switch_to.window(remaining[0])
            else:
                raise
        time.sleep(1)

    print(f"=== Loco {loco_no} complete ===")


# ============================================================================
# MAIN
# ============================================================================

def main():
    df = load_sheet_data()

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)  # browser khula rahe script khatam hone ke baad
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.get(LOGIN_URL)
    print("\n[ACTION REQUIRED] Browser mein khud login karein (username, password, captcha).")
    print("Login + 'Add Loco Equipment' page tak pahunchne ke baad yahan Terminal mein Enter dabayein...")
    input()

    unique_locos = df["Loco No"].dropna().unique()
    unique_locos = [l for l in unique_locos if not is_blank(l)]
    print(f"\n[INFO] Sheet mein {len(unique_locos)} unique Loco Numbers mile: {list(unique_locos)}")

    for loco_no in unique_locos:
        try:
            process_loco(driver, df, loco_no)
        except Exception as e:
            print(f"[ERROR] Loco {loco_no} process karte waqt issue: {e}")
            continue

    print("\n[DONE] Saare locos process ho gaye. Browser check kar lein ki sab sahi save hua.")


if __name__ == "__main__":
    main()
