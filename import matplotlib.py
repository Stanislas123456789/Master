import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt

# Définir les cryptos et les prix d'achat
cryptos = ['BTC-USD', 'ETH-USD']
purchase_prices = {'BTC-USD': 25000, 'ETH-USD': 3000}

# Télécharger les données actuelles
data = yf.download(cryptos, period='1d', interval='1d')['Adj Close']

# Obtenir les prix actuels
current_prices = data.iloc[-1]

# Calculer la performance
performance = {crypto: (current_prices[crypto] - purchase_prices[crypto]) / purchase_prices[crypto] * 100 for crypto in cryptos}

# Afficher les résultats dans le terminal
print("Prix actuels :")
print(current_prices)
print("\nPerformance des investissements :")
print(performance)

# Créer un graphique pour les prix actuels
plt.figure(figsize=(10, 5))
plt.bar(current_prices.index, current_prices.values, color='blue', alpha=0.7)
plt.title('Prix Actuels des Cryptos')
plt.xlabel('Crypto')
plt.ylabel('Prix en USD')
plt.savefig('crypto_prices.png')  # Sauvegarder le graphique
plt.show()

# Créer un graphique pour les performances
plt.figure(figsize=(10, 5))
plt.bar(performance.keys(), performance.values(), color='green', alpha=0.7)
plt.title('Performance des Investissements (%)')
plt.xlabel('Crypto')
plt.ylabel('Performance (%)')
plt.savefig('crypto_performance.png')  # Sauvegarder le graphique
plt.show()
