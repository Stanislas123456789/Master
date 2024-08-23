from flask import Flask
import dash
from dash import dcc, html
import dash_table
import yfinance as yf
import threading
import webbrowser

# Configuration de Flask
server = Flask(__name__)

# Configuration de Dash
app = dash.Dash(__name__, server=server, url_base_pathname='/dashboard/')

# Fonction pour ouvrir automatiquement le navigateur
def open_browser():
    url = 'http://127.0.0.1:5000/dashboard/'
    threading.Timer(1.25, lambda: webbrowser.open(url)).start()

# Fonction pour obtenir les données des cryptomonnaies
def fetch_crypto_data():
    tickers = ['BTC-USD', 'ETH-USD']
    data = {}
    for ticker in tickers:
        crypto = yf.Ticker(ticker)
        info = crypto.info
        data[ticker] = {
            'Nom': ticker,
            'Prix': info.get('currentPrice', 'N/A'),
            '% 24h': info.get('regularMarketChangePercent', 'N/A'),
            '7d %': 'N/A',  # Placeholder, change as needed
            '30d %': 'N/A', # Placeholder, change as needed
            '90d %': 'N/A', # Placeholder, change as needed
            'Cap. Marché': info.get('marketCap', 'N/A')
        }
    return data

# Obtention des données
crypto_data = fetch_crypto_data()

# Création de la disposition du dashboard
app.layout = html.Div(children=[
    html.H1(children='Crypto Dashboard', style={'textAlign': 'center'}),

    html.Div(className='card', children=[
        dash_table.DataTable(
            id='crypto-table',
            columns=[{'name': col, 'id': col} for col in crypto_data[list(crypto_data.keys())[0]].keys()],
            data=[value for value in crypto_data.values()],
            style_table={'overflowX': 'auto'},
            style_header={'backgroundColor': '#007bff', 'color': 'white', 'fontWeight': 'bold'},
            style_data={'backgroundColor': '#f9f9f9', 'color': '#333'},
            style_data_conditional=[
                {
                    'if': {'row_index': 'odd'},
                    'backgroundColor': '#ffffff'
                }
            ]
        )
    ], style={'padding': '20px', 'margin': '0 auto', 'width': '90%'})
])

# Route pour l'index
@server.route('/')
def index():
    return 'Bienvenue sur le Dashboard de Crypto! Cliquez <a href="/dashboard/">ici</a> pour voir le tableau de bord.'

if __name__ == '__main__':
    open_browser()
    server.run(debug=True, port=5000)

