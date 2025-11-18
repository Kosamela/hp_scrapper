from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import openpyxl
import time
import re
import os

HP_SEARCH = "https://support.hp.com/us-en/products?jumpid=in_r11839_us/en/WIN11_MDA"


def scrap_with_selenium(serials):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    results = []

    for serial in serials:
        print(f"\n➡ Pobieram dane dla: {serial}")
        driver.get(HP_SEARCH)


        # --- Klikamy Accept Cookies ---
        try:
            accept_btn = driver.find_element(By.ID, "onetrust-accept-btn-handler")
            accept_btn.click()
            print("✅ Kliknięto Accept Cookies")
            time.sleep(1)
        except:
            print("ℹ️ Brak popupu cookies lub już zaakceptowane")

        # --- Wyszukujemy numer seryjny ---
        try:
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "searchQueryField"))
            )
            search_box.clear()
            search_box.send_keys(serial)
            search_box.send_keys(Keys.ENTER)
            print(f"✅ Wpisano numer seryjny: {serial}")
        except:
            print("❌ Nie znaleziono pola wyszukiwania")
            continue

        # --- Czekamy aż strona produktu się załaduje ---
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "productSpecContainer"))
            )
            print("✅ Strona produktu załadowana")
        except:
            print("❌ Timeout: strona produktu się nie załadowała")
            continue

        # --- Kliknięcie 'View full specifications', jeśli jest ---
        try:
            full_spec_btn = driver.find_element(By.ID, "Viewfull")
            full_spec_btn.click()
            print("✅ Kliknięto View full specifications")
            time.sleep(2)
        except:
            print("ℹ️ Brak przycisku View full specifications")

        # --- Parsowanie BeautifulSoup ---
        soup = BeautifulSoup(driver.page_source, "html.parser")
        data = {"Model": None, "Serial": serial, "CPU": None, "RAM": None, "Dysk": None, "OS": None, "Gwarancja": None}

        # --- Pobranie modelu ---
        model_tag = soup.find("h1", class_="product-name-text")
        if model_tag:
            data["Model"] = model_tag.text.strip()
            print(f"✅ Model: {data['Model']}")

        # --- Warranty ---
        warr = soup.find("div", class_="common")
        if warr and "Warranty status" in warr.text:
            if "expired" in warr.text.lower():
                data["Gwarancja"] = "Expired"
            else:
                data["Gwarancja"] = "Active"
            print(f"✅ Gwarancja: {data['Gwarancja']}")

        # --- Specyfikacje produktu ---
        specs = soup.find("div", id="productSpecContainer")
        if specs:
            spec_items = specs.find_all("div", class_="spec-content")
            for item in spec_items:
                title = item.find("div", class_="spec-title")
                value = item.find("div", class_="desc-text-non-view-encapsulation")

                if not value:
                    app_desc = item.find("app-description-text-product-spec")
                    if app_desc:
                        value = app_desc.find("div", class_="desc-text-non-view-encapsulation")

                t = title.text.strip().lower() if title else ""
                v = value.text.strip() if value else ""

                # --- Skracanie danych ---
                if "operating system" in t:
                    os_match = re.match(r"(Windows\s\d+\s\w+)", v)
                    data["OS"] = os_match.group(0) if os_match else v
                    print(f"✅ OS: {data['OS']}")
                elif t == "processor":
                    cpu_match = re.match(r"(Intel® Core™\s*[iI]\d{1,2}-\d{3,4}[A-Za-z]*)", v)
                    data["CPU"] = cpu_match.group(0).strip() if cpu_match else v
                    print(f"✅ CPU: {data['CPU']}")
                elif "memory" in t or "ram" in t:
                    if not data["RAM"]:
                        ram_match = re.search(r"\d+\s*GB", v)
                        if ram_match:
                            data["RAM"] = ram_match.group(0).strip()
                            print(f"✅ RAM: {data['RAM']}")
                elif "storage" in t or "hard drive" in t or "ssd" in t or "internal drive" in t:
                    if not data["Dysk"] and v:
                        size_match = re.search(r"(\d+\s*(?:GB|TB))", v, re.IGNORECASE)
                        type_match = "SSD" if "ssd" in v.lower() else "HDD"
                        if size_match:
                            data["Dysk"] = f"{size_match.group(1).upper()} {type_match}"
                            print(f"✅ Dysk: {data['Dysk']}")

        results.append(data)

    driver.quit()
    return results


def save_excel(data, file="hp_laptops.xlsx"):
    if os.path.exists(file):
        os.remove(file)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Model", "Serial", "CPU", "RAM", "Dysk", "OS", "Gwarancja"])

    for row in data:
        ws.append([row["Model"], row["Serial"], row["CPU"], row["RAM"], row["Dysk"], row["OS"], row["Gwarancja"]])

    wb.save(file)
    print(f"📁 Wynik zapisany do: {file}")


def main():
    serials = ["SN", "SN"]  # Tutaj kurla wpisz ten sn w postaci listy z przecinkiem np jak kolejne ,"SN"
    results = scrap_with_selenium(serials)
    save_excel(results)


if __name__ == "__main__":
    main()

