import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Configureer de WebDriver
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)

# Laad de bestaande Excel met de URLs
data_df = pd.read_excel('4personenURL.xlsx')  # Pas de naam aan indien nodig

# Lijsten voor de verzamelde punten
scullpunten = []
boordpunten = []

# Loop door de data en verzamel Scull en Boord punten
for index, row in data_df.iterrows():
    url = row['URL']  # Zorg ervoor dat je de juiste kolomnaam gebruikt
    full_url = url  # Aangezien de volledige URL nu in de data staat

    print(f"Probeer te openen: {full_url}")  # Debugging output
    try:
        driver.get(full_url)  # Volledige URL openen
        time.sleep(3)  # Wacht tot de pagina is geladen

        # Zoek de Scull punten
        scull_div = driver.find_element(By.XPATH, "//div[contains(text(), 'Scull:')]/following-sibling::div")
        scullpunten.append(scull_div.text.strip())  # Voeg het gevonden getal toe

        # Zoek de Boord punten
        boord_div = driver.find_element(By.XPATH, "//div[contains(text(), 'Boord:')]/following-sibling::div")
        boordpunten.append(boord_div.text.strip())  # Voeg het gevonden getal toe

    except Exception as e:
        print(f'Fout bij het ophalen van gegevens voor {full_url}: {e}')
        scullpunten.append(None)  # Geen data gevonden
        boordpunten.append(None)  # Geen data gevonden

# Voeg de punten toe aan de DataFrame
data_df['Scullpunten'] = scullpunten
data_df['Boordpunten'] = boordpunten

# Sla de bijgewerkte DataFrame op in een nieuw Excel-bestand
data_df.to_excel('5PersonenPunten.xlsx', index=False)

print("Data verzameld en opgeslagen in 'personen_resultaten_met_punten.xlsx'.")
driver.quit()
