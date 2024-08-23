import dash
from dash import dcc, html, dash_table
import pandas as pd
import threading
import webbrowser
from flask import Flask

# Fonction pour lire les données depuis un fichier Excel et les transformer en tableau
def read_data_from_excel(file_path):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    print(f"Feuilles disponibles : {sheet_names}")

    # Remplacer 'Invest' par le nom correct de la feuille si nécessaire
    sheet_name = 'Invest'  # Assurez-vous que ce nom est correct

    if sheet_name not in sheet_names:
        raise ValueError(f"Feuille nommée '{sheet_name}' non trouvée dans le fichier Excel.")

    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Calculer la valeur de position
    df['Valeur de Position'] = df['NB TOKEN'] * df['PRICE']

    # Assurez-vous que les colonnes sont correctement nommées et formatées
    df = df.round({'PRIX D\'achat': 2, 'PRICE': 2, 'Valeur de Position': 2, '%': 2, '7d %': 2, '30d %': 2, '90d %': 2, 'MCAP ATH': 2, 'MCAP Target': 2})
    df['MCAP ATH'] = df['MCAP ATH'].apply(lambda x: f"${x:,.2f}")
    df['MCAP Target'] = df['MCAP Target'].apply(lambda x: f"${x:,.2f}")
    df['Valeur de Position'] = df['Valeur de Position'].apply(lambda x: f"${x:,.2f}")
    
    return df.to_dict('records')

# Configuration de Flask
server = Flask(__name__)

# Configuration de Dash
app = dash.Dash(__name__, server=server, url_base_pathname='/dashboard/')

# Chemin vers le fichier Excel (corrigé)
file_path = r'C:\Users\Smichel\Desktop\MASTER\templates\Invest.xlsx'

# Fonction pour ouvrir automatiquement le navigateur
def open_browser():
    url = 'http://127.0.0.1:5000/dashboard/'
    threading.Timer(1.25, lambda: webbrowser.open(url)).start()

# Layout du tableau
app.layout = html.Div([
    html.Div(className='card', children=[
        html.H1(children='Crypto Dashboard'),

        dash_table.DataTable(
            id='crypto-table',
            columns=[
                {"name": "Nom", "id": "Coinmarketcap"},  # ou "Coingecko" si nécessaire
                {"name": "PRIX D'achat", "id": "PRIX D'achat"},
                {"name": "Nombre de Crypto", "id": "NB TOKEN"},
                {"name": "Prix Actuel", "id": "PRICE"},
                {"name": "Valeur de Position", "id": "Valeur de Position"},
                {"name": "% 24h", "id": "%"},
                {"name": "7d %", "id": "7d %"},
                {"name": "30d %", "id": "30d %"},
                {"name": "90d %", "id": "90d %"},
                {"name": "Cap. Marché ATH", "id": "MCAP ATH"},
                {"name": "Cap. Marché Target", "id": "MCAP Target"}
            ],
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'right', 'padding': '5px'},
            style_header={'backgroundColor': '#007bff', 'color': 'white'},
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9f9f9'}
            ],
            data=read_data_from_excel(file_path)
        )
    ]),

    dcc.Interval(
        id='interval-component',
        interval=30*1000,  # Intervalle en millisecondes (30 secondes)
        n_intervals=0
    )
])

# Callback pour mettre à jour les données du tableau
@app.callback(
    dash.dependencies.Output('crypto-table', 'data'),
    [dash.dependencies.Input('interval-component', 'n_intervals')]
)
def update_table(n):
    return read_data_from_excel(file_path)

# Route pour l'index
@server.route('/')
def index():
    return 'Bienvenue sur le Dashboard de Crypto! Cliquez <a href="/dashboard/">ici</a> pour voir le tableau de bord.'

if __name__ == '__main__':
    open_browser()
    server.run(debug=True, port=5000)
