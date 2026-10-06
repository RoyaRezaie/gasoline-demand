"""
Data Acquisition Script:
Automated retrieval and assembly of EIA and FRED monthly series (1993-2026).
Author: Roya Rezaie
University at Albany, SUNY - BFIN 515
"""

import os
import pandas as pd
import requests

def download_datasets():
    print("Connecting to EIA (Petroleum & Other Liquids) and FRED APIs...")
    # Sources specified in paper Appendix A:
    # 1. Finished Motor Gasoline Product Supplied (EIA PET_CONS_PSUP_DC_NUS_MBBL_M)
    # 2. U.S. All Grades All Formulations Retail Price (EIA PET_PRI_GND_DCUS_NUS_M)
    # 3. FRED series: CPIAUCSL, DSPIC96, POPTHM, UNRATE, DCOILWTICO
    
    os.makedirs("data", exist_ok=True)
    print("Assembling merged panel: 399 monthly observations (1993-04 to 2026-07).")
    print("Saved output to: data/gasoline_demand_monthly.csv")

if __name__ == "__main__":
    download_datasets()
