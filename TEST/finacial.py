import pandas as pd

# List of financial terms
financial_terms = [
    "actors", "market participants", "limit order book", "bid-ask", "liquidity", "stocks", "dividends", "repo", "bonds", 
    "interest rates", "yield curve", "zero-coupon curve", "duration", "convexity", "term structure", "swap rates", "swaps", 
    "commodities", "convenience yield", "carry cost", "carry", "rolldown", "contango", "backwardation", "futures", 
    "margining", "non-arbitrage conditions", "replication", "credit", "spread", "options", "greeks", "risk", 
    "volatilities", "volatility surface"
]

# Corresponding explanations for each financial term
explanations = [
    "Individuals or entities participating in financial markets.",
    "Participants who engage in the buying and selling of financial instruments.",
    "A record of all outstanding orders to buy and sell a particular financial instrument.",
    "The difference between the highest price a buyer is willing to pay and the lowest price a seller will accept.",
    "The ease with which an asset can be bought or sold in the market without affecting its price.",
    "Securities representing ownership in a corporation and a claim on part of its assets and earnings.",
    "Payments made by a corporation to its shareholders, usually as a distribution of profits.",
    "A form of short-term borrowing for dealers in government securities.",
    "Debt instruments issued by corporations or governments to raise capital.",
    "The cost of borrowing money, usually expressed as a percentage.",
    "A graph showing the relationship between bond yields and maturities.",
    "A curve that plots the yields of zero-coupon bonds against their maturities.",
    "A measure of the sensitivity of the price of a bond to a change in interest rates.",
    "A measure of the curvature of the relationship between bond prices and bond yields.",
    "The relationship between interest rates or bond yields and different terms or maturities.",
    "The interest rates applicable to swap contracts.",
    "Financial contracts in which two parties exchange streams of payments over time.",
    "Raw materials or primary agricultural products that can be bought and sold.",
    "The benefit or premium associated with holding a physical commodity over a futures contract.",
    "The costs associated with holding a financial position, including interest and storage costs.",
    "The net return on an investment after accounting for the costs of carrying the position.",
    "The change in the yield of a bond or futures contract as it approaches maturity.",
    "A situation where the futures price is higher than the expected future spot price.",
    "A situation where the futures price is lower than the expected future spot price.",
    "Financial contracts obligating the buyer to purchase an asset or the seller to sell an asset at a predetermined future date and price.",
    "The process of settling the profits and losses of futures contracts on a daily basis.",
    "Conditions where no arbitrage opportunities exist, ensuring fair pricing in the market.",
    "The process of mimicking the payoff of a financial instrument using a combination of other instruments.",
    "The risk of loss due to a borrower's failure to repay a loan or meet contractual obligations.",
    "The difference between the yield on a corporate bond and a government bond of similar maturity.",
    "Financial derivatives that provide the right, but not the obligation, to buy or sell an asset at a specified price.",
    "Measures of sensitivity of the price of options to changes in underlying parameters.",
    "The possibility of losing money on an investment or business venture.",
    "Statistical measures of the dispersion of returns for a given security or market index.",
    "A three-dimensional plot showing implied volatilities for different option strikes and expirations."
]

# Create a DataFrame
df = pd.DataFrame({
    "Financial Terms": financial_terms,
    "Explanations": explanations
})

# Save the DataFrame to an Excel file
file_path = "financial_terms_with_explanations.xlsx"
df.to_excel(file_path, index=False)

print(f"Excel file has been generated: {file_path}")
