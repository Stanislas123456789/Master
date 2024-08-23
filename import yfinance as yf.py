import yfinance as yf

# Définir les cryptos et les prix d'achat
cryptos = ['BTC-USD', 'ETH-USD']
purchase_prices = {'BTC-USD': 25000, 'ETH-USD': 3000}

# Télécharger les données actuelles
data = yf.download(cryptos, period='1d', interval='1d')['Adj Close']

# Obtenir les prix actuels
current_prices = data.iloc[-1]

# Calculer la performance
performance = {crypto: (current_prices[crypto] - purchase_prices[crypto]) / purchase_prices[crypto] * 100 for crypto in cryptos}

# Afficher les résultats
print("Prix actuels :")
print(current_prices)
print("\nPerformance des investissements :")
print(performance)

