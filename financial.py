import pandas as pd
import numpy as np
import numpy_financial as npf  

def calculate_financials():
    # Initial Investment Breakdown (from Table 2 in document)
    initial_investments = {
        'Buildings': 820_000_000,
        'Limestone Equipment': 451_185_381,
        'Staff Salaries (Year 1)': 643_000_000,
        'Laboratory Equipment': 156_639_720,
        'Lime Granulation Equipment': 841_803_796,
        'Training': 60_000_000,
        'Utilities Setup': 184_500_000,
        'Laboratory Installation': 20_136_005,
        'Factory Installation': 34_716_250,
        'Packaging (6 months)': 874_800_000,
        'Fertilizer Plates': 30_000_000,
        'Market Research': 214_000_000,
        'Contingency (5%)': 216_539_058,
    }

    # Calculate total initial investment
    total_investment = sum(initial_investments.values())

    # Production Capacity Calculations (from document)
    daily_production = 600  # tons per day
    lime_price_per_kg = 103  # RWF
    working_days_per_year = 323.62  # from document

    # Calculate annual production and revenue
    annual_production = daily_production * working_days_per_year
    annual_revenue = annual_production * 1000 * lime_price_per_kg  # converting tons to kg

    # Operating Costs (estimated from provided data)
    annual_operating_costs = {
        'Staff Salaries': 643_000_000,  # from document
        'Utilities': 24_000_000,  # from utilities setup table
        'Packaging Materials': 874_800_000 * 2,  # assuming 6-month cost * 2
        'Maintenance': total_investment * 0.02,  # estimated at 2% of initial investment
        'Marketing': 214_000_000,  # from document
        'Other Operating Expenses': 60_180_000,  # from inventory costs
    }

    total_annual_opex = sum(annual_operating_costs.values())

    # Calculate annual cash flows
    annual_cash_flow = annual_revenue - total_annual_opex

    # Create cash flow projections for 10 years
    discount_rate = 0.13  # 13% as specified
    cash_flows = [-total_investment]  # Year 0 is initial investment
    cash_flows.extend([annual_cash_flow] * 10)  # Append 10 years of cash flows

    # Calculate NPV
    npv = npf.npv(discount_rate, cash_flows)

    # Calculate IRR
    irr = npf.irr(cash_flows)

    # Calculate Payback Period
    cumulative_cash_flow = np.cumsum(cash_flows)
    payback_period_years = np.where(cumulative_cash_flow >= 0)[0][0]  # First year where cumulative cash flow is positive

    # Create results DataFrame
    results_df = pd.DataFrame({
        'Metric': ['Initial Investment (RWF)', 'Annual Revenue (RWF)', 'Annual Operating Costs (RWF)',
                   'Net Annual Cash Flow (RWF)', 'NPV (RWF)', 'IRR', 'Payback Period (Years)'],
        'Value': [total_investment, annual_revenue, total_annual_opex,
                  annual_cash_flow, npv, irr, payback_period_years]
    })

    # Create cash flow DataFrame
    cash_flow_df = pd.DataFrame({
        'Year': range(11),
        'Cash Flow (RWF)': cash_flows,
        'Cumulative Cash Flow (RWF)': cumulative_cash_flow
    })

    return results_df, cash_flow_df

# Calculate and format results
results_df, cash_flow_df = calculate_financials()

# Save results to Excel
results_df.to_excel('lime_plant_financial_analysis.xlsx', index=False)
cash_flow_df.to_excel('lime_plant_cash_flows.xlsx', index=False)

print("Financial analysis and cash flow data exported to Excel.")
