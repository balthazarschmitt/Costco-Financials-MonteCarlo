# Import packages
import openpyxl as pyxl
import numpy as np
import yfinance as yf

# Load the excel workbook and define the two relevant sheets
wb = pyxl.load_workbook('Costco Financial Model.xlsx', data_only=True)
ws1 = wb['3_Statement_Model']
ws2 = wb['DCF']

# Extract relevant inputs from the workbook
revenue_0 = ws1['H6'].value
tax_rate = ws1['I23'].value
net_debt = ws2['N80'].value
shares = ws2['N82'].value
base_cash = ws1['H37'].value + ws1['H38'].value

# Define total number of simulations
N = 20000

# Generate random variables for the Monte Carlo simulation as Normal distributions
growth_1_5 = np.random.normal(0.06, 0.015, N)
growth_6_10 = np.random.normal(0.04, 0.01, N)
ebit_margin = np.random.normal(0.037, 0.003, N)

# % of revenue
dep_pct = np.random.normal(0.009, 0.005, N) # Depreciation
capex_pct = np.random.normal(0.019, 0.007, N) # CapEx
delta_nwc_pct = np.random.normal(0.002, 0.0007, N) # NWC

# Generate random variables for wacc and tgr as a Normal and Triangular distribution respectively
wacc = np.random.normal(0.0734, 0.005, N)
tgr = np.random.triangular(0.025, 0.03, 0.035, N)

# Clamp depreciation, capex, and nwc percentages to within reasonable bounds
dep_pct = np.clip(dep_pct, 0.0, 0.10)
capex_pct = np.clip(capex_pct, 0.0, 0.15)
delta_nwc_pct = np.clip(delta_nwc_pct, 0.0, 0.01)

# Clamp wacc and tgr to within reasonable bounds
wacc = np.clip(wacc, 0.05, 0.12)
tgr = np.clip(tgr, 0.0, 0.04)

# Define main DCF function that runs within each simulation
def dcf_value(revenue_0, g1, g2, margin, tax_rate, wacc, tgr, dep_ratio, capex_ratio, nwc_ratio):
    # Project revenues (10 years)
    revenues = []
    revenue = revenue_0
    for t in range(1, 11):
        g = g1 if t <= 5 else g2
        revenue *= (1 + g)
        revenues.append(revenue)

    rev = np.array(revenues)

    # EBIT and NOPAT
    ebit = rev * margin
    nopat = ebit * (1 - tax_rate)

    # D&A, CapEx, NWC levels per year (use scalar ratios for this simulation)
    dep = rev * dep_ratio
    capex = rev * capex_ratio
    delta_nwc = rev * nwc_ratio

    # FCFF per year
    fcff = nopat + dep - capex - delta_nwc
    
    # Cap FCFF to be non-negative
    fcff[-1] = max(fcff[-1], 0)

    # Discount and terminal value
    discount_factors = np.array([(1 + wacc) ** t for t in range(1, 11)])
    pv_fcff = np.sum(fcff / discount_factors)

    denom = max(wacc - tgr, 1e-6)  # guard against wacc <= tgr
    terminal_value = fcff[-1] * (1 + tgr) / denom
    pv_terminal = terminal_value / ((1 + wacc) ** 10)

    enterprise_value = pv_fcff + pv_terminal
    equity_value = enterprise_value - net_debt + base_cash
    return equity_value

# Run the monte carlo simulation using the defined DCF function above N times
equity_values = np.array([
    dcf_value(
        revenue_0,
        growth_1_5[i],
        growth_6_10[i],
        ebit_margin[i],
        tax_rate,
        wacc[i],
        tgr[i],
        dep_pct[i],      # pass scalar ratios for this simulation
        capex_pct[i],
        delta_nwc_pct[i]
    )
    for i in range(N)
])

# Calculate the value per share for each total equity value in the list
value_per_share = equity_values / shares

# Fetch current stock price for Costco
ticker = yf.Ticker("COST")
hist = ticker.history(period="1d")
current_price = hist['Close'].iloc[-1]

print("Monte Carlo Simulation Result:\n")

print("KEY INSIGHTS:\n")

print("Value per share (USD):")
print("5th Percentile:    ", np.percentile(value_per_share, 5).round(2))
print("25th Percentile:   ", np.percentile(value_per_share, 25).round(2))
print("50th Percentile:   ", np.percentile(value_per_share, 50).round(2))
print("75th Percentile:   ", np.percentile(value_per_share, 75).round(2))
print("95th Percentile:   ", np.percentile(value_per_share, 95).round(2))

print("\nKey Statistics:")
print("Mean:              ", np.mean(value_per_share).round(2))
print("Standard Deviation:", np.std(value_per_share).round(2))
print("Median:            ", np.median(value_per_share).round(2))
print("IQR:               ", (np.percentile(value_per_share, 75) - np.percentile(value_per_share, 25)).round(2))
print("Minimum:           ", np.min(value_per_share).round(2))
print("Maximum:           ", np.max(value_per_share).round(2))

print("\nOTHER INSIGHTS:\n")

print("Return Potential (%):")
print("5th Percentile:    ", ((np.percentile(value_per_share, 5).round(2) / current_price - 1)*100).round(2), "%")
print("25th Percentile:   ", ((np.percentile(value_per_share, 25).round(2) / current_price - 1)*100).round(2), "%")
print("50th Percentile:   ", ((np.percentile(value_per_share, 50).round(2) / current_price - 1)*100).round(2), "%")
print("75th Percentile:   ", ((np.percentile(value_per_share, 75).round(2) / current_price - 1)*100).round(2), "%")
print("95th Percentile:   ", ((np.percentile(value_per_share, 95).round(2) / current_price - 1)*100).round(2), "%")

print("\nPercent of Simulations Above Current Price:")

above = value_per_share > current_price
pct_above = np.mean(above) * 100
pct_below = 100 - pct_above   
print(f"Above:              {pct_above:.2f}%       ({np.count_nonzero(above)}/{N}) simulations")
print(f"Below:              {pct_below:.2f}%      ({N - np.count_nonzero(above)}/{N}) simulations")

