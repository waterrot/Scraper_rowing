import pandas as pd

# Lees de data van de twee Excel-bestanden
database_df = pd.read_excel('database_met_url.xlsx')  # Dit bestand bevat Naam, URL, Club
nieuwe_personen_df = pd.read_excel('nieuwe_personen_data.xlsx')  # Dit bestand bevat Naam, Club

# Maak een lege lijst voor de resultaten
resultaten = []

# Loop door elke rij in het nieuwe personen DataFrame
for index, row in nieuwe_personen_df.iterrows():
    naam = row['Naam']
    club = row['Club']
    
    # Zoek naar een match in de database DataFrame
    match = database_df[(database_df['Naam'] == naam) & (database_df['Club'].str.contains(club, na=False))]
    
    # Voeg de originele gegevens en de URL (indien beschikbaar) toe aan de resultaten
    if not match.empty:
        url = match['URL'].values[0]  # Neem de eerste match
    else:
        url = None  # Geen match gevonden

    # Voeg de gegevens toe aan de resultatenlijst
    resultaten.append({
        "Naam": naam,
        "Club": club,
        "URL": url
    })

# Maak een DataFrame van de resultaten
resultaten_df = pd.DataFrame(resultaten)

# Sla de resultaten op in een nieuw Excel-bestand
resultaten_df.to_excel('vergelijking_resultaten.xlsx', index=False)

print("Vergelijking voltooid en resultaten opgeslagen in 'vergelijking_resultaten.xlsx'.")
