# Costco Monte Carlo DCF Valuation

A probabilistic discounted cash flow (DCF) model for Costco Wholesale Corporation (COST) that uses Monte Carlo simulation to quantify valuation uncertainty across 20,000 scenarios.

## Overview

Rather than relying on a single point estimate, this model generates a distribution of intrinsic values by varying key assumptions: revenue growth rates, operating margins, capital expenditure requirements, and discount rates. The output reveals not just an expected value, but the full range of probable outcomes given realistic parameter uncertainty.

## Methodology

The simulation draws inputs from the following distributions:

- **Revenue Growth**: Years 1-5 at 6% ± 1.5%, years 6-10 at 4% ± 1% (normal)
- **EBIT Margin**: 3.7% ± 0.3% (normal)
- **WACC**: 7.34% ± 0.5% (normal)
- **Terminal Growth Rate**: 2.5-3.5% (triangular)
- **CapEx/Revenue**: 1.9% ± 0.7% (normal, clamped 0-15%)
- **Depreciation/Revenue**: 0.9% ± 0.5% (normal, clamped 0-10%)
- **∆NWC/Revenue**: 0.2% ± 0.07% (normal, clamped 0-1%)

Each simulation projects 10 years of free cash flow to the firm (FCFF), calculates terminal value, and converts enterprise value to equity value per share.

## Key Results

Based on current analysis with **market price of $967.08**:

| Percentile | Value/Share | Implied Return |
|------------|-------------|----------------|
| 5th        | $44.52      | -95.4%         |
| 25th       | $189.29     | -80.4%         |
| 50th       | $295.58     | -69.5%         |
| 75th       | $407.47     | -57.9%         |
| 95th       | $576.89     | -40.4%         |

**Mean**: $302.63 | **Std Dev**: $159.88 | **Range**: -$72.94 to $1,216.96

**Critical Finding**: 99.97% of simulations (19,994/20,000) produce valuations below the current market price, suggesting either significant overvaluation or structural issues in the model assumptions.

## Requirements

```bash
pip install openpyxl numpy yfinance matplotlib
```

## Usage

Ensure `Costco Financial Model.xlsx` is in the working directory with populated sheets:
- `3_Statement_Model`: Historical financials and base assumptions
- `DCF`: Net debt, shares outstanding, WACC components

Run the simulation:

```bash
python montecarlo.py
```

Output includes percentile statistics, return potential analysis, and a histogram visualization comparing the distribution against current market price.

## Interpretation

The model's bearish signal warrants investigation into several potential explanations:

1. **Assumption Calibration**: EBIT margins (3.7%), CapEx ratios (1.9%), and depreciation rates (0.9%) may not reflect Costco's actual economics, particularly the high-margin membership fee model
2. **Quality Premium**: Market may be pricing intangible factors like brand strength, operational moat, and consistent execution that a pure DCF doesn't capture
3. **Growth Expectations**: 6% near-term growth may be conservative for a retailer successfully expanding internationally
4. **Base Revenue**: Verify the starting revenue figure aligns with Costco's most recent fiscal year (~$254B FY2024)

## License

CC0 1.0 Universal