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

# Generate random variables for wacc and tgr as a Normal and Triangular distribution respectively
wacc = np.random.normal(0.0734, 0.005, N)
tgr = np.random.triangular(0.025, 0.03, 0.035, N)

# Clamp wacc and tgr to within reasonable bounds
wacc = np.clip(wacc, 0.05, 0.12)
tgr = np.clip(tgr, 0.0, 0.04)

# Define main DCF function that runs within each simulation
def dcf_value(revenue_0, g1, g2, margin, tax_rate, wacc, tgr):
    # Create empty list for projected revenues
    revenues = []
    # Create temp variable for current rev in each year starting at the initial rev
    revenue = revenue_0
    
    # Loop through the 10 forcasted years
    for t in range(1, 11):
        # Use growth rate 1 or 2 based on year (first 1-5 vs 6-10)
        g = g1 if t <= 5 else g2
        # Update rev for the current year using the define growth rate
        revenue *= (1 + g)
        # Add the each individual growth rate to the rev list
        revenues.append(revenue)
        
    # Calculate EBIT for each year based off of the projected revenues and the ebit margin defined before
    ebit = np.array(revenues) * margin
    
    # Calculate NOPAT by removing taxes from EBIT
    nopat = ebit * (1 - tax_rate)
    
    # Approximate FCFF as NOPAT (ignoring changes in working capital and capex for simplicity)
    fcff = nopat
    
    # Calculate the discount factors for each year
    discount_factors = np.array([(1 + wacc)**t for t in range(1, 11)])
    # Calculate PV of FCFFs using the discount factors
    pv_fcff = np.sum(fcff / discount_factors)
    
    # Calculate terminal value using the Gordon Growth Model
    terminal_value = fcff[-1] * (1 + tgr) / (wacc - tgr)
    # Calculate PV of terminal value
    pv_terminal = terminal_value / ((1 + wacc) ** 10)
    
    # Calculate total enterprise value
    enterprise_value = pv_fcff + pv_terminal
    # Calculate total equity value by subtracting net debt from enterprise value
    equity_value = enterprise_value - net_debt + base_cash
    
    # Return the calculated total equity value
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
        tgr[i]
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

