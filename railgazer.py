import os
import time
from io import StringIO

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.utils import get_column_letter


FOIS_URL = "https://www.fois.indianrail.gov.in/FOISWebPortal/pages/FWP_RASIOSttnWiseOtsgDmndN.jsp"
STATION_FROM_CODES = ["BYFS", "ISCG", "FOS", "SOBK", "PBSB", "IISM", "HLSR", "SSMK"]
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "fois_station_from_8codes.xlsx")

def get_driver():
    options = webdriver.ChromeOptions()

    
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
   
    options.add_argument("--disable-blink-features=AutomationControlled")

    
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver



def set_ckp_all_all(driver, timeout=30):
    print("SDTD welcomes you to RailGazer!\n")
    wait = WebDriverWait(driver, timeout)

    def find_input(field_name):
        
        try:
            return wait.until(EC.presence_of_element_located(
                (By.NAME, field_name)
            ))
        except Exception:
            return wait.until(EC.presence_of_element_located(
                (By.XPATH, f"//input[@name='{field_name}' or @id='{field_name}']")
            ))

    
    txt_div = find_input("txtDvsn")
    
    txt_clst = find_input("txtClst")
    
    txt_stn = find_input("txtSttn")

    
    for elem, val, label in [
        (txt_div, "CKP", "Division"),
        (txt_clst, "ALL", "Cluster"),
        (txt_stn, "ALL", "Station"),
    ]:
        try:
            elem.clear()
        except Exception:
            pass
        elem.send_keys(val)
    

    time.sleep(1)


def click_submit(driver, timeout=30):
    
    wait = WebDriverWait(driver, timeout)

    
    try:
        btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 "//button[contains(translate(normalize-space(), "
                 "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'SUBMIT')]")
            )
        )
        
        btn.click()
        return
    except Exception:
        print("Error: UI changes or page not loaded. Please contact Team SDTD.")

    
    btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH,
             "//input[(@type='button' or @type='submit') and "
             "contains(translate(@value, 'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'SUBMIT')]")
        )
    )
    
    btn.click()


def find_tablesorter_html(driver, timeout=60):
    
    end_time = time.time() + timeout
    tablesorter_xpath = "//table[contains(@class, 'tablesorter')]"

    while time.time() < end_time:
        driver.switch_to.default_content()

        
        tables = driver.find_elements(By.XPATH, tablesorter_xpath)
        if tables:
            
            return tables[0].get_attribute("outerHTML")

        
        frames = driver.find_elements(By.TAG_NAME, "iframe") + \
                 driver.find_elements(By.TAG_NAME, "frame")
        
            

        for idx, frame in enumerate(frames):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(frame)
                tables = driver.find_elements(By.XPATH, tablesorter_xpath)
                if tables:
                    print(f"[*] Found tablesorter table inside frame {idx}.")
                    html = tables[0].get_attribute("outerHTML")
                    driver.switch_to.default_content()
                    return html
            except Exception as e:
                print(f"    Frame {idx} error: {e!r}")
                driver.switch_to.default_content()
                continue

        time.sleep(1)

    driver.switch_to.default_content()
    raise RuntimeError("Timeout: could not find <table> with class containing 'tablesorter'.")




def extract_main_table(html: str) -> pd.DataFrame:
    
    tables = pd.read_html(StringIO(html))
    

    if not tables:
        raise RuntimeError("No tables found in HTML.")

    target_df = None
    for i, df in enumerate(tables):
        cols_upper = [str(c).strip().upper() for c in df.columns]
    
        if any("STATION FROM" in c for c in cols_upper):
    
            target_df = df
            break

    if target_df is None:
    
        target_df = tables[0]

    
    return target_df


def filter_by_station_from(df: pd.DataFrame, code: str) -> pd.DataFrame:
    code = code.upper().strip()

    col_map = {str(c): str(c).upper().strip() for c in df.columns}
    station_col = None
    for orig, upper in col_map.items():
        if "STATION FROM" in upper:
            station_col = orig
            break

    if station_col is None:
        raise RuntimeError("Could not find 'STATION FROM' column in table.")

    df = df.copy()
    df[station_col] = df[station_col].astype(str).str.upper().str.strip()
    filtered = df[df[station_col] == code].copy()
    print(f"    [{code}] rows:", len(filtered))
    return filtered


def autosize_columns_for_sheet(ws, df: pd.DataFrame):
    
    for i, col in enumerate(df.columns, start=1):
        column_letter = get_column_letter(i)
        try:
            max_len = max(
                [len(str(col))]
                + [len(str(v)) for v in df[col].astype(str).tolist()]
            )
        except ValueError:
            max_len = len(str(col))
        ws.column_dimensions[column_letter].width = max_len + 2



def main():
    driver = get_driver()
    try:
        
        driver.get(FOIS_URL)
        time.sleep(3)

        
        set_ckp_all_all(driver)

        
        click_submit(driver)

        
        table_html = find_tablesorter_html(driver)
        debug_path = os.path.join(SCRIPT_DIR, "fois_table_debug.html")
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(table_html)
        
        df_full = extract_main_table(table_html)


        with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
            for code in STATION_FROM_CODES:
                
                df_code = filter_by_station_from(df_full, code)
                sheet_name = code[:31]
                df_code.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                autosize_columns_for_sheet(ws, df_code)



    except Exception as e:
        print("ERROR:", repr(e))
    finally:
        print("Thank you for using RailGazer - Team SDTD")
        time.sleep(5)
        driver.quit()


if __name__ == "__main__":
    main()
