from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Instellingen voor de browser
options = webdriver.ChromeOptions()

# Start Chrome WebDriver
driver = webdriver.Chrome(options=options)

# Ga naar de website
driver.get('https://hoesnelwasik.nl/tbr/0/loting/11ef-8a66-4acccfaa-bfa6-fa53f5d3a545')

# Wacht even zodat je de pagina kunt laden
time.sleep(3)

# Maak een lege lijst om alle data in op te slaan
data = []
headers = []  # Lijst om headers op te slaan

# Zoek alle rijen van de hoofd-tabel
rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')

# Loop door alle rijen en verzamel de data
for row in rows:
    try:
        # Sluit de pop-up indien deze open is
        try:
            close_button = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.modal-footer button.close'))
            )
            close_button.click()
        except:
            pass  # Geen pop-up om te sluiten

        # Zoek de eerste kolom (het logo) en klik om de pop-up te openen
        logo = row.find_element(By.CSS_SELECTOR, 'td:first-child img')
        driver.execute_script("arguments[0].click();", logo)

        # Wacht totdat de modaal verschijnt (maximaal 10 seconden)
        popup_modal = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.modal-body'))
        )

        # Wacht tot de tabel in de modal-body volledig is geladen
        popup_table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.modal-body table'))
        )

        # Haal de data uit de modale tabel
        popup_rows = popup_table.find_elements(By.CSS_SELECTOR, 'tr')

        # Maak een lijst voor de rijen uit de pop-up
        popup_data = []

        for popup_row in popup_rows:
            # Haal de tekst uit elke cel in de pop-up tabel
            columns = popup_row.find_elements(By.CSS_SELECTOR, 'td')
            row_data = []

            for col in columns:
                div_element = col.find_element(By.TAG_NAME, 'div') if col.find_elements(By.TAG_NAME, 'div') else None
                if div_element:
                    # Haal de tekst en data-label op
                    text = div_element.text.strip()
                    data_label = col.get_attribute('data-label').strip()
                    row_data.append(text)

                    # Voeg data-label toe aan headers als deze nog niet bestaat
                    if data_label not in headers:
                        headers.append(data_label)

            if row_data:  # Controleer of er data is
                popup_data.append(row_data)

        # Voeg de verzamelde data van de pop-up toe aan de data
        # Voeg de data uit de pop-up toe als nieuwe rijen in de data-lijst
        for data_row in popup_data:
            data.append(data_row)

        # Klik op de achtergrond om de pop-up te sluiten
        body = driver.find_element(By.CSS_SELECTOR, 'body')
        body.click()

        # Wacht even om de browser niet te overbelasten
        time.sleep(2)

    except Exception as e:
        print(f"Fout opgetreden bij verwerken van een rij: {e}")
        continue

# Sluit de browser
driver.quit()

# Zet de data in een pandas DataFrame
df = pd.DataFrame(data, columns=headers)

# Exporteer de DataFrame naar een Excel-bestand
df.to_excel('scraped_data.xlsx', index=False)

print("Data succesvol opgeslagen in 'scraped_data.xlsx'")
