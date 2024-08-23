import pandas as pd

# Chemins vers les fichiers Excel
chemin_fichier_existing = r'H:\DIRECTION\Clients\MyMoneyBank (ex-GE CAPITAL BANK- SOVAC)\2023 11 Projet Dauphin\02 - Contrat\Contrats Flux\0. Post approval\1. Analyse\Comparaison_V2_V3IO.xlsx'
chemin_fichier_resultats = r'H:\DIRECTION\Clients\MyMoneyBank (ex-GE CAPITAL BANK- SOVAC)\2023 11 Projet Dauphin\02 - Contrat\Contrats Flux\0. Post approval\1. Analyse\Comparaison_Consolidation.xlsx'

# Charger les fichiers Excel dans des DataFrames
df_V2 = pd.read_excel(chemin_fichier_existing, sheet_name='Consolidation V2')
df_V3 = pd.read_excel(chemin_fichier_existing, sheet_name='Consolidation V3')

# Assurez-vous que 'account_number' est le bon nom ou ajustez le nom si nécessaire
id_column = ' account_number'  # Nom de la colonne d'identifiant

# Convertir les colonnes de dates en datetime si nécessaire
date_columns = [' dt_entree_ca']  # Ajuster les colonnes de date si nécessaire
for col in date_columns:
    if col in df_V2.columns:
        df_V2[col] = pd.to_datetime(df_V2[col], errors='coerce', dayfirst=True)
    if col in df_V3.columns:
        df_V3[col] = pd.to_datetime(df_V3[col], errors='coerce', dayfirst=True)

# Réduire les DataFrames aux colonnes communes
common_columns = df_V2.columns.intersection(df_V3.columns).tolist()
df_V2 = df_V2[common_columns]
df_V3 = df_V3[common_columns]

# Vérifier la présence de la colonne d'identifiant
if id_column not in df_V2.columns or id_column not in df_V3.columns:
    raise KeyError(f"La colonne '{id_column}' n'est pas présente dans les DataFrames.")

# Identifier les dossiers ajoutés, enlevés et modifiés
df_V2_dossiers = df_V2.set_index(id_column)
df_V3_dossiers = df_V3.set_index(id_column)

# Dossiers ajoutés
dossiers_ajoutes = df_V3_dossiers[~df_V3_dossiers.index.isin(df_V2_dossiers.index)]

# Dossiers enlevés
dossiers_enleves = df_V2_dossiers[~df_V2_dossiers.index.isin(df_V3_dossiers.index)]

# Dossiers modifiés
dossiers_communs = df_V2_dossiers[df_V2_dossiers.index.isin(df_V3_dossiers.index)]
dossiers_modifies = dossiers_communs[dossiers_communs.ne(df_V3_dossiers.loc[dossiers_communs.index]).any(axis=1)]

# Filtrer pour conserver seulement les dossiers où 'encours_datatape' a changé
if 'encours_datatape' in dossiers_modifies.columns:
    dossiers_modifies = dossiers_modifies[
        dossiers_modifies['encours_datatape'] != df_V3_dossiers.loc[dossiers_modifies.index, 'encours_datatape']
    ]

# Créer un nouveau fichier Excel pour les résultats
with pd.ExcelWriter(chemin_fichier_resultats, engine='xlsxwriter') as writer:
    # Écrire les onglets de consolidation
    df_V2.to_excel(writer, sheet_name='Consolidation V2', index=False)
    df_V3.to_excel(writer, sheet_name='Consolidation V3', index=False)
    
    # Écrire les dossiers ajoutés
    dossiers_ajoutes.reset_index().to_excel(writer, sheet_name='Dossiers Ajoutés', index=False)

    # Écrire les dossiers enlevés
    dossiers_enleves.reset_index().to_excel(writer, sheet_name='Dossiers Enlevés', index=False)

    # Écrire les dossiers modifiés
    dossiers_modifies.reset_index().to_excel(writer, sheet_name='Dossiers Modifiés', index=False)
    
    # Ajouter un onglet récapitulatif
    summary_data = {
        'Onglet': ['Dossiers Ajoutés', 'Dossiers Enlevés', 'Dossiers Modifiés'],
        'Nombre de Dossiers': [len(dossiers_ajoutes), len(dossiers_enleves), len(dossiers_modifies)],
        'Somme Encours': [
            dossiers_ajoutes['encours_datatape'].sum() if 'encours_datatape' in dossiers_ajoutes.columns else 'N/A',
            dossiers_enleves['encours_datatape'].sum() if 'encours_datatape' in dossiers_enleves.columns else 'N/A',
            dossiers_modifies['encours_datatape'].sum() if 'encours_datatape' in dossiers_modifies.columns else 'N/A'
        ]
    }
    df_summary = pd.DataFrame(summary_data)
    df_summary.to_excel(writer, sheet_name='Récapitulatif', index=False)

    # Optionnel : formatage des onglets ajoutés
    workbook = writer.book

    bold_format = workbook.add_format({'bold': True, 'align': 'center'})
    center_format = workbook.add_format({'align': 'center'})

    def format_sheet(sheet_name):
        worksheet = workbook.get_worksheet_by_name(sheet_name)
        header_row = 0

        # Appliquer le formatage pour les en-têtes
        for col_num, value in enumerate(dossiers_ajoutes.columns):
            worksheet.write(header_row, col_num + 1, value, bold_format)  # Décaler le titre de 1 cellule

        # Définir la largeur des colonnes
        for col_num, value in enumerate(dossiers_ajoutes.columns):
            max_length = max(dossiers_ajoutes[value].astype(str).map(len).max(), len(value))
            worksheet.set_column(col_num + 1, col_num + 1, max_length, center_format)  # Décaler la largeur de 1 colonne

    # Formater chaque feuille
    for sheet_name in ['Dossiers Ajoutés', 'Dossiers Enlevés', 'Dossiers Modifiés', 'Récapitulatif']:
        format_sheet(sheet_name)
    
    # Ajouter un commentaire pour les dossiers modifiés
    modified_sheet = workbook.get_worksheet_by_name('Dossiers Modifiés')
    modified_sheet.write(0, len(dossiers_modifies.columns) + 2, 'Commentaires', bold_format)
    for i in range(1, len(dossiers_modifies) + 1):
        modified_sheet.write(i, len(dossiers_modifies.columns) + 2, 'Le dossier a été modifié.', center_format)