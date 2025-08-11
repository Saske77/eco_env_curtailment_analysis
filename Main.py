import pandas as pd
from datetime import datetime, timedelta
import os
import warnings

warnings.filterwarnings('ignore')

# Configuration Parameters
TURBINE_CAPACITY_MW = 2.3
COMPENSATION_RATE = 0.925
KWK1_ANLAGENSCHLUESSEL = 'E20793019000000000000004288000004'

# File paths (assuming files are in 'data' folder)
DATA_FOLDER = 'data'
FILES = {
    'curtailment': 'export-finished-2025-06-26 13_40_41.xlsx',  # Fixed filename
    'market': 'Gro_handelspreise_202401010000_202501020000_Stunde.xlsx',
    'redispatch': 'Ausgleichsenergie_202401010000_202501020000_Stunde.xlsx',
    'strommix': 'stromix_2024.xlsx'
}


def load_curtailment_data(filepath):
    """Load and process curtailment data"""
    print(f"Loading curtailment data from {filepath}")

    try:
        df = pd.read_excel(filepath, sheet_name=0)
        df.columns = df.columns.str.strip()

        df['Start_DT'] = pd.to_datetime(df['Start'], errors='coerce')
        df['End_DT'] = pd.to_datetime(df['Ende'], errors='coerce')

        date_failures = df[df['Start_DT'].isna() | df['End_DT'].isna()]
        if not date_failures.empty:
            print(f"Warning: Failed to parse dates for {len(date_failures)} records")
            print(date_failures[['Start', 'Ende']].head())

        df = df.dropna(subset=['Start_DT', 'End_DT'])

        invalid_dates = df[df['End_DT'] < df['Start_DT']]
        if not invalid_dates.empty:
            df = df[df['End_DT'] >= df['Start_DT']]

        # Filter for 2024 and KWK1
        df_2024 = df[df['Start_DT'].dt.year == 2024].copy()
        print(f"Records in 2024: {len(df_2024)}")

        # Filter for specific plant
        df_kwk1 = df_2024[df_2024['Anlagenschlüssel'] == KWK1_ANLAGENSCHLUESSEL].copy()
        print(f"Filtered for plant {KWK1_ANLAGENSCHLUESSEL}: {len(df_kwk1)} events")

        if len(df_kwk1) == 0:
            print("Warning: No data found for the specified plant and year")
            return pd.DataFrame()

        # Process curtailment level and duration
        df_kwk1['Curtailment_Level'] = pd.to_numeric(df_kwk1['Stufe (%)'], errors='coerce').fillna(0)
        df_kwk1['Duration_Min'] = pd.to_numeric(df_kwk1['Dauer (Min)'], errors='coerce')

        # Calculate duration from start/end times (in minutes)
        df_kwk1['Calculated_Duration'] = (df_kwk1['End_DT'] - df_kwk1['Start_DT']).dt.total_seconds() / 60

        # Use calculated duration where provided duration is missing or invalid
        mask = df_kwk1['Duration_Min'].isna() | (df_kwk1['Duration_Min'] <= 0)
        df_kwk1.loc[mask, 'Duration_Min'] = df_kwk1.loc[mask, 'Calculated_Duration']

        # Final cleanup
        df_kwk1 = df_kwk1[df_kwk1['Duration_Min'] > 0]
        df_kwk1 = df_kwk1.drop(columns=['Calculated_Duration'])

        print(f"Found {len(df_kwk1)} curtailment events for KWK1 in 2024")
        print(f"Total curtailment duration: {df_kwk1['Duration_Min'].sum():.0f} minutes")
        print(f"Max event duration: {df_kwk1['Duration_Min'].max():.0f} minutes")
        print(f"Min event duration: {df_kwk1['Duration_Min'].min():.0f} minutes")
        print(f"Average curtailment level: {df_kwk1['Curtailment_Level'].mean():.1f}%")

        return df_kwk1

    except Exception as e:
        print(f"Error loading curtailment data: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


def load_market_data(filepath):
    """Load and process market price data"""
    print(f"Loading market data from {filepath}")

    try:
        # Read Excel file, skipping header rows
        df = pd.read_excel(filepath, sheet_name='Großhandelspreise', skiprows=9)
        print(f"Market data shape: {df.shape}")

        # Clean column names
        df.columns = df.columns.str.strip()

        # Convert datetime with German format (DD.MM.YYYY HH:MM)
        df['DateTime'] = pd.to_datetime(df['Datum von'], format='%d.%m.%Y %H:%M', errors='coerce')

        # Convert price, handling German number format (comma as decimal separator)
        price_col = 'Deutschland/Luxemburg [€/MWh]'
        df['Price'] = pd.to_numeric(
            df[price_col].astype(str).str.replace(',', '.'),
            errors='coerce'
        )

        # Remove invalid rows
        df = df.dropna(subset=['DateTime', 'Price'])

        # Set datetime as index
        df = df.set_index('DateTime')

        print(f"Loaded {len(df)} market price records")
        print(f"Price range: {df['Price'].min():.2f} to {df['Price'].max():.2f} €/MWh")
        print(f"Average price: {df['Price'].mean():.2f} €/MWh")

        return df[['Price']]

    except Exception as e:
        print(f"Error loading market data: {e}")
        return pd.DataFrame()


def load_redispatch_data(filepath):
    """Load and process redispatch data"""
    print(f"Loading redispatch data from {filepath}")

    try:
        # Read Excel file, skipping header rows
        df = pd.read_excel(filepath, sheet_name='Ausgleichsenergie', skiprows=9)
        print(f"Redispatch data shape: {df.shape}")

        # Clean column names
        df.columns = df.columns.str.strip()

        # Convert datetime with German format
        df['DateTime'] = pd.to_datetime(df['Datum von'], format='%d.%m.%Y %H:%M', errors='coerce')

        # Convert price, handling German number format
        df['Price'] = pd.to_numeric(
            df['Preis [€/MWh]'].astype(str).str.replace(',', '.'),
            errors='coerce'
        )

        # Remove invalid rows
        df = df.dropna(subset=['DateTime', 'Price'])

        # Set datetime as index
        df = df.set_index('DateTime')

        print(f"Loaded {len(df)} redispatch price records")
        print(f"Price range: {df['Price'].min():.2f} to {df['Price'].max():.2f} €/MWh")
        print(f"Average price: {df['Price'].mean():.2f} €/MWh")

        return df[['Price']]

    except Exception as e:
        print(f"Error loading redispatch data: {e}")
        return pd.DataFrame()


def load_strommix_data(filepath):
    """Load and process CO2 emission factor data"""
    print(f"Loading strommix data from {filepath}")

    try:
        # Read Excel file
        df = pd.read_excel(filepath, sheet_name='DE_2024_hourly')
        print(f"Strommix data shape: {df.shape}")

        # Clean column names
        df.columns = df.columns.str.strip()

        # The CO2 column has special characters, find it by partial match
        co2_col = None
        for col in df.columns:
            if 'Carbon intensity' in col and 'direct' in col:
                co2_col = col
                break

        if co2_col is None:
            print("Error: Could not find CO2 intensity column")
            return pd.DataFrame()

        # Convert datetime from UTC format
        df['DateTime'] = pd.to_datetime(df['Datetime (UTC)'], errors='coerce')

        # Convert CO2 factor
        df['CO2_Factor'] = pd.to_numeric(df[co2_col], errors='coerce')

        # Remove invalid rows
        df = df.dropna(subset=['DateTime', 'CO2_Factor'])

        # Convert UTC to local time (assuming German time zone)
        df['DateTime'] = df['DateTime'].dt.tz_localize('UTC').dt.tz_convert('Europe/Berlin').dt.tz_localize(None)

        # Set datetime as index
        df = df.set_index('DateTime')

        print(f"Loaded {len(df)} CO2 emission factor records")
        print(f"CO2 factor range: {df['CO2_Factor'].min():.1f} to {df['CO2_Factor'].max():.1f} g/kWh")
        print(f"Average CO2 factor: {df['CO2_Factor'].mean():.1f} g/kWh")

        return df[['CO2_Factor']]

    except Exception as e:
        print(f"Error loading strommix data: {e}")
        return pd.DataFrame()


def calculate_curtailment_impact(curtailment_df, market_df, redispatch_df, strommix_df):
    """Calculate economic and environmental impact of curtailment"""

    if curtailment_df.empty:
        print("No curtailment events found")
        return None

    total_curtailed_energy = 0
    total_missed_revenue = 0
    total_redispatch_cost = 0
    total_co2_emissions = 0
    processed_events = 0

    print(f"Analyzing {len(curtailment_df)} curtailment events...")

    for idx, event in curtailment_df.iterrows():
        start_time = event['Start_DT']
        end_time = event['End_DT']
        curtailment_level = event['Curtailment_Level']

        # Calculate curtailment percentage
        # Assuming 'Stufe (%)' represents the remaining capacity percentage
        # So curtailment = 100% - remaining capacity
        curtailment_percent = (100 - curtailment_level) / 100 if curtailment_level < 100 else 0

        if curtailment_percent <= 0:
            continue  # Skip if no actual curtailment

        processed_events += 1

        # Calculate hourly impacts
        current_time = start_time
        while current_time < end_time:
            # Find next hour boundary
            next_hour = current_time.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
            segment_end = min(next_hour, end_time)

            # Duration of this segment in hours
            duration_hours = (segment_end - current_time).total_seconds() / 3600

            # Curtailed energy in this segment
            curtailed_energy = duration_hours * TURBINE_CAPACITY_MW * curtailment_percent
            total_curtailed_energy += curtailed_energy

            # Get hourly data (using hour start time)
            hour_key = current_time.floor('H')

            # Market price
            market_price = 0
            if not market_df.empty and hour_key in market_df.index:
                market_price = market_df.loc[hour_key, 'Price']

            # Redispatch price
            redispatch_price = 0
            if not redispatch_df.empty and hour_key in redispatch_df.index:
                redispatch_price = redispatch_df.loc[hour_key, 'Price']

            # CO2 factor (g/kWh)
            co2_factor = 0
            if not strommix_df.empty and hour_key in strommix_df.index:
                co2_factor = strommix_df.loc[hour_key, 'CO2_Factor']

            # Calculate impacts
            missed_revenue = curtailed_energy * market_price
            redispatch_cost = curtailed_energy * redispatch_price
            co2_emissions = curtailed_energy * 1000 * co2_factor / 1000  # Convert to kg

            total_missed_revenue += missed_revenue
            total_redispatch_cost += redispatch_cost
            total_co2_emissions += co2_emissions

            current_time = segment_end

    # Calculate final results
    compensation_paid = total_missed_revenue * COMPENSATION_RATE
    total_economic_impact = compensation_paid + total_redispatch_cost
    total_co2_tonnes = total_co2_emissions / 1000

    return {
        'Total Curtailed Energy (MWh)': total_curtailed_energy,
        'Total Missed Revenue (€)': total_missed_revenue,
        'Compensation Paid (€)': compensation_paid,
        'Total Redispatch Cost (€)': total_redispatch_cost,
        'Total Economic Impact (€)': total_economic_impact,
        'Total CO2 Emissions (tonnes)': total_co2_tonnes,
        'Number of Events': len(curtailment_df),
        'Processed Events (with curtailment)': processed_events,
        'Total Duration (minutes)': curtailment_df['Duration_Min'].sum()
    }


def main():
    """Main function to run the analysis"""
    print("=== Wind Turbine Curtailment Analysis ===\n")

    # Load all data files
    curtailment_df = load_curtailment_data(os.path.join(DATA_FOLDER, FILES['curtailment']))

    if curtailment_df.empty:
        print("Cannot proceed without curtailment data.")
        return

    market_df = load_market_data(os.path.join(DATA_FOLDER, FILES['market']))
    redispatch_df = load_redispatch_data(os.path.join(DATA_FOLDER, FILES['redispatch']))
    strommix_df = load_strommix_data(os.path.join(DATA_FOLDER, FILES['strommix']))

    print("\n" + "=" * 50)

    # Calculate impact
    results = calculate_curtailment_impact(curtailment_df, market_df, redispatch_df, strommix_df)

    if results:
        print("\n=== CURTAILMENT IMPACT ANALYSIS RESULTS ===")
        print(f"Analysis Period: 2024")
        print(f"Wind Turbine: KWK 1 (Capacity: {TURBINE_CAPACITY_MW} MW)")
        print(f"Compensation Rate: {COMPENSATION_RATE * 100}%")
        print("-" * 50)

        for key, value in results.items():
            if isinstance(value, float):
                print(f"{key}: {value:,.2f}")
            else:
                print(f"{key}: {value:,}")

        # Additional insights
        print("\n=== KEY INSIGHTS ===")
        if results['Total Curtailed Energy (MWh)'] > 0:
            avg_curtailment_per_event = results['Total Curtailed Energy (MWh)'] / results[
                'Processed Events (with curtailment)']
            print(f"Average curtailed energy per event: {avg_curtailment_per_event:.2f} MWh")

            capacity_factor_loss = (results['Total Curtailed Energy (MWh)'] / (TURBINE_CAPACITY_MW * 8760)) * 100
            print(f"Annual capacity factor loss due to curtailment: {capacity_factor_loss:.2f}%")

            avg_price = results['Total Missed Revenue (€)'] / results['Total Curtailed Energy (MWh)'] if results[
                                                                                                             'Total Curtailed Energy (MWh)'] > 0 else 0
            print(f"Average electricity price during curtailment: {avg_price:.2f} €/MWh")

    else:
        print("Analysis could not be completed due to data issues.")


if __name__ == "__main__":
    main()