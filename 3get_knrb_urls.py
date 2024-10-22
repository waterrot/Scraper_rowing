import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Stel de ChromeDriver in
chrome_driver_path = 'chromedriver.exe'  # Vervang dit door het pad naar jouw chromedriver
service = Service(chrome_driver_path)

# Configureer Chrome opties om SSL fouten te negeren
options = webdriver.ChromeOptions()
options.add_argument("--ignore-certificate-errors")

# Start de webdriver
driver = webdriver.Chrome(service=service, options=options)

# Lees de nieuwe personen data
personen_df = pd.read_excel('nieuwe_personen_data.xlsx')

# Voor het opslaan van resultaten
resultaten = []

# Loop door elke naam in de eerste kolom (kolom A)
for index, row in personen_df.iterrows():
    naam = row.iloc[0]  # Gebruik iloc om toegang te krijgen tot de eerste kolom (kolom A)

    # Ga naar de zoekpagina
    driver.get("https://roeievenementen.knrb.nl/person-search")

    # Vul de naam in het zoekveld in
    try:
        zoekveld = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "__BVID__12"))  # ID van het zoekveld
        )
        zoekveld.clear()
        zoekveld.send_keys(naam)

        # Klik op de zoekknop
        zoek_knop = driver.find_element(By.CSS_SELECTOR, "button.btn-primary")  # Pas de selector aan indien nodig
        zoek_knop.click()

        time.sleep(3)  # Wacht tot de resultaten geladen zijn

        # Zoek naar de resultaten in de tabel
        personen_lijst = driver.find_elements(By.CSS_SELECTOR, "table#__BVID__14 tbody tr")  # Correcte selector voor de tbody van de tabel

        # Loop door de gevonden resultaten
        for persoon in personen_lijst:
            try:
                # Verkrijg naam en href van het <a> element
                naam_element = persoon.find_element(By.TAG_NAME, "a")
                naam_text = naam_element.text.strip()  # Verwijder spaties
                href_value = naam_element.get_attribute("href")

                # Verkrijg de club uit de tweede kolom (td)
                club_text = persoon.find_elements(By.TAG_NAME, "td")[1].text.strip()  # Tweede td voor club

                # Voeg de gegevens toe aan de resultaten
                resultaten.append({"Naam": naam_text, "URL": href_value, "Club": club_text})
            except Exception as e:
                print(f'Fout bij het ophalen van gegevens voor {naam}: {e}')

    except Exception as e:
        print(f'Fout bij het verwerken van {naam}: {e}')

# Maak een DataFrame van de resultaten
resultaten_df = pd.DataFrame(resultaten)

# Sla de resultaten op in een nieuw Excel-bestand
resultaten_df.to_excel('database_met_url.xlsx', index=False)

print("Gegevens succesvol verzameld en opgeslagen in 'personen_resultaten.xlsx'.")

# Sluit de webdriver
driver.quit()
