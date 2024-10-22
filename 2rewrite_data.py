import pandas as pd

# Lees de gegevens uit de eerder opgeslagen Excel
df = pd.read_excel('scraped_data.xlsx')

# Maak een lege lijst voor de nieuwe data
new_data = []

# Loop door de rijen in de DataFrame
for index, row in df.iterrows():
    organisatie = row['Club']  # Organisatie in kolom A

    # Loop door de naam kolommen van E t/m H (Boeg, 2, 3, Slag)
    for col in ['Boeg', '2', '3', 'Slag']:
        naam = row[col]  # Naam in de respectieve kolom
        new_data.append([naam, organisatie, '', ''])  # Voeg de naam, organisatie, en lege waarden toe

# Zet de nieuwe data in een DataFrame
new_df = pd.DataFrame(new_data, columns=['Naam', 'Club', 'Boord', 'Scull'])

# Exporteer de DataFrame naar een nieuwe Excel-bestand
new_df.to_excel('nieuwe_personen_data.xlsx', index=False)

print("Nieuwe data succesvol opgeslagen in 'nieuwe_personen_data.xlsx'")
