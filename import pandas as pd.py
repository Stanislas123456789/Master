import pandas as pd

# Données des programmes financiers
data = [
    {"Rank": 1, "School": "London Business School (LBS) - Masters in Financial Analysis", "City": "London", "Country": "United Kingdom", "Tuition Fees (Approx.)": "£37,900", "Tuition Fees (EUR)": "~€44,343", "QS Ranking": 2, "FT Ranking": 3, "Economist Ranking": 1},
    {"Rank": 2, "School": "University of Oxford - Saïd Business School - MSc in Financial Economics", "City": "Oxford", "Country": "United Kingdom", "Tuition Fees (Approx.)": "£51,490", "Tuition Fees (EUR)": "~€60,243", "QS Ranking": 1, "FT Ranking": 2, "Economist Ranking": 3},
    {"Rank": 3, "School": "University of Cambridge - Judge Business School - Master of Finance (MFin)", "City": "Cambridge", "Country": "United Kingdom", "Tuition Fees (Approx.)": "£48,000", "Tuition Fees (EUR)": "~€56,160", "QS Ranking": 3, "FT Ranking": 1, "Economist Ranking": 2},
    # Ajoutez d'autres lignes ici
]

# Création d'un DataFrame Pandas
df = pd.DataFrame(data)

# Affichage du DataFrame
print(df)

# Chemin vers votre fichier Excel de sortie
output_excel_file = 'finance_programs_rankings.xlsx'

# Exporter le DataFrame vers un fichier Excel
df.to_excel(output_excel_file, index=False)

print(f"Données exportées avec succès vers '{output_excel_file}'")



