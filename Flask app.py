from flask import Flask, render_template
import yfinance as yf
import webbrowser
import os
import threading

app = Flask(__name__)

def open_browser():
    url = 'http://127.0.0.1:5000'
    # Utilise le chemin correct vers chrome.exe
    chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
    os.startfile(f'{chrome_path} {url}')
    
# Démarre le navigateur dans un thread séparé pour ne pas bloquer le serveur Flask
threading.Thread(target=open_browser).start()

@app.route('/')
def index():
    cryptos = ['BTC-USD', 'ETH-USD']
    purchase_prices = {'BTC-USD': 25000, 'ETH-USD': 3000}

    data = yf.download(cryptos, period='1d', interval='1d')['Adj Close']
    current_prices = data.iloc[-1]
    performance = {crypto: (current_prices[crypto] - purchase_prices[crypto]) / purchase_prices[crypto] * 100 for crypto in cryptos}

    return render_template('index.html', current_prices=current_prices, performance=performance)

if __name__ == '__main__':
    app.run(debug=True)
