"""
Replication Script for:
The Impact of Gasoline Prices on Gasoline Demand in the United States (1993-2026)
Author: Roya Rezaie
University at Albany, SUNY - BFIN 515: Economic Analysis
"""

import numpy as np
import pandas as pd
import statsmodels.api as sm
from statsmodels.sandbox.regression.gmm import IV2SLS

def load_data(filepath="gasoline_demand_monthly.csv"):
    """
    Loads and prepares monthly EIA and FRED data (1993-2026).
    Expected series: quantity (finished motor gasoline supplied),
    price_real (2017$), income_per_capita, and wti_real.
    """
    df = pd.read_csv(filepath)
    df['date'] = pd.to_datetime(df['date'])
    df = df.sort_values('date').reset_index(drop=True)
    
    # Natural logarithmic transforms
    df['ln_Q'] = np.log(df['quantity'])
    df['ln_P'] = np.log(df['price_real'])
    df['ln_Y'] = np.log(df['income_per_capita'])
    if 'wti_real' in df.columns:
        df['ln_WTI'] = np.log(df['wti_real'])
        
    # Time trend (t=0 at April 1993)
    df['trend'] = np.arange(len(df))
    
    # 2007 Structural break terms
    break_date = pd.to_datetime('2007-01-01')
    df['step_2007'] = (df['date'] >= break_date).astype(int)
    df['trend_post2007'] = np.where(
        df['date'] >= break_date,
        (df['trend'] - df.loc[df['date'] == break_date, 'trend'].values[0]) / 12.0,
        0.0
    )
    
    # Pandemic mobility restriction dummy (March 2020 - June 2021)
    df['pandemic'] = ((df['date'] >= '2020-03-01') & (df['date'] <= '2021-06-30')).astype(int)
    
    # Monthly seasonal indicators (December base)
    df['month'] = df['date'].dt.month
    for m in range(1, 12):
        df[f'm_{m}'] = (df['month'] == m).astype(int)
        
    return df

def run_models(df):
    month_cols = [f'm_{m}' for m in range(1, 12)]
    
    print("==================================================")
    print("Model 1: Baseline OLS (Linear Trend)")
    print("==================================================")
    X1 = sm.add_constant(df[['ln_P', 'ln_Y', 'trend'] + month_cols])
    m1 = sm.OLS(df['ln_Q'], X1).fit(cov_type='HAC', cov_kwds={'maxlags': 12})
    print(f"Price Elasticity: {m1.params['ln_P']:.4f} (HAC s.e.: {m1.bse['ln_P']:.4f})")
    
    print("\n==================================================")
    print("Model 2: Preferred OLS (2007 Break + Pandemic Dummy)")
    print("==================================================")
    X2_cols = ['ln_P', 'ln_Y', 'trend', 'step_2007', 'trend_post2007', 'pandemic'] + month_cols
    X2 = sm.add_constant(df[X2_cols])
    m2 = sm.OLS(df['ln_Q'], X2).fit(cov_type='HAC', cov_kwds={'maxlags': 12})
    print(f"Preferred Price Elasticity: {m2.params['ln_P']:.4f} (HAC s.e.: {m2.bse['ln_P']:.4f})")
    print(f"R-squared: {m2.rsquared:.4f}")
    
    print("\n==================================================")
    print("Model 3: ARDL Dynamic Adjustment Specification")
    print("==================================================")
    df['ln_Q_lag'] = df['ln_Q'].shift(1)
    df_ardl = df.dropna().copy()
    X3 = sm.add_constant(df_ardl[['ln_P', 'ln_Q_lag', 'ln_Y', 'trend', 'step_2007', 'trend_post2007', 'pandemic'] + month_cols])
    m3 = sm.OLS(df_ardl['ln_Q'], X3).fit(cov_type='HAC', cov_kwds={'maxlags': 12})
    b = m3.params['ln_P']
    rho = m3.params['ln_Q_lag']
    long_run = b / (1.0 - rho)
    print(f"Short-Run Elasticity: {b:.4f} (HAC s.e.: {m3.bse['ln_P']:.4f})")
    print(f"Adjustment Coefficient (rho): {rho:.4f}")
    print(f"Implied Long-Run Elasticity: {long_run:.4f}")
    
    return m1, m2, m3

if __name__ == "__main__":
    print("EIA/FRED Gasoline Demand Econometric Model Replication Script")
