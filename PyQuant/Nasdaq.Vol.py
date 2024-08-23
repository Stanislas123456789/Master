import yfinance as yf
import numpy as np
from datetime import datetime, timedelta

# Définir la période des 20 dernières années
end_date = datetime.today()
start_date = end_date - timedelta(days=20*365)  # Environ 20 ans

# Liste des indices à analyser
indices = {
    'NASDAQ-100': '^NDX',
    'S&P 500': '^GSPC',
    'Russell 2000': '^RUT'
}

# Taux sans risque (supposons 2%)
risk_free_rate = 0.02

for name, ticker in indices.items():
    # Télécharger les données historiques
    index = yf.Ticker(ticker)
    index_hist = index.history(start=start_date, end=end_date)
    
    # Calculer les rendements journaliers
    index_hist['Daily Return'] = index_hist['Close'].pct_change()
    
    # Calculer l'écart type annuel
    std_dev_annual = np.std(index_hist['Daily Return'].dropna()) * np.sqrt(252)
    
    # Calculer le rendement moyen annuel
    mean_daily_return = np.mean(index_hist['Daily Return'].dropna())
    mean_annual_return = mean_daily_return * 252
    
    # Calculer le ratio de Sharpe
    sharpe_ratio = (mean_annual_return - risk_free_rate) / std_dev_annual
    
    # Afficher les résultats
    print(f"Pour l'indice {name} ({ticker}):")
    print(f" - L'écart type annuel sur les 20 dernières années est de : {std_dev_annual}")
    print(f" - Le ratio de Sharpe sur les 20 dernières années est de : {sharpe_ratio}")
    print(f" - Le Rendement est de : {mean_annual_return}")
    print()  # Ligne vide pour séparer les résultats
