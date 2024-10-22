import pandas as pd

# Laad de data van de twee Excel-bestanden
scraped_data_df = pd.read_excel('1scraped_data.xlsx')  # Originele data
punten_df = pd.read_excel('5PersonenPunten.xlsx')  # Gegevens met punten

# Maak kolommen voor totaalScull en totaalBoord
scraped_data_df['totaalScull'] = 0
scraped_data_df['totaalBoord'] = 0

# Loop door elke rij in de scraped_data_df DataFrame
for index, row in scraped_data_df.iterrows():
    # Voor elke roeier in de kolommen E t/m H
    for col in ['Boeg', '2', '3', 'Slag']:  # Pas aan naar de juiste kolomnamen indien nodig
        roeier = row[col]
        if pd.notna(roeier):  # Controleer of de roeier niet leeg is
            # Zoek de overeenkomstige rij in de punten DataFrame
            match = punten_df[punten_df['Naam'].str.contains(roeier, na=False)]  # Zorg dat de naam overeenkomt

            # Als er overeenkomsten zijn, tel de punten op
            if not match.empty:
                row['totaalScull'] += match['Scullpunten'].sum()  # Telt de scullpunten bij elkaar op
                row['totaalBoord'] += match['Boordpunten'].sum()  # Telt de boordpunten bij elkaar op

    # Update de DataFrame met de nieuwe totaalwaarden
    scraped_data_df.at[index, 'totaalScull'] = row['totaalScull']
    scraped_data_df.at[index, 'totaalBoord'] = row['totaalBoord']

# Sla de bijgewerkte DataFrame op in een nieuw Excel-bestand met meerdere sheets
with pd.ExcelWriter('6data_van_alles.xlsx') as writer:
    scraped_data_df.to_excel(writer, sheet_name='Resultaten', index=False)  # Hoofdsheet met resultaten
    punten_df.to_excel(writer, sheet_name='5PersonenPunten', index=False)  # Extra sheet met punten

print("Data verzameld en opgeslagen in '6data_van_alles.xlsx' met een extra sheet voor '5PersonenPunten'.")
