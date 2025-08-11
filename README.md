# Wind Turbine Curtailment Analysis

A Python tool for analyzing the economic and environmental impact of wind turbine curtailment events using German energy market data.

## Overview

This tool analyzes curtailment events for wind turbines, calculating the economic losses and environmental impacts when turbines are forced to reduce their power output due to grid constraints. The analysis combines curtailment data with market prices, redispatch costs, and CO2 emission factors to provide comprehensive impact assessment.

## Features

- **Economic Impact Analysis**: Calculate missed revenue and compensation payments
- **Redispatch Cost Assessment**: Analyze grid balancing costs associated with curtailment
- **Environmental Impact**: Calculate CO2 emissions from curtailed renewable energy
- **Detailed Reporting**: Generate comprehensive reports with key insights
- **Data Validation**: Built-in diagnostics for data quality assessment

## Requirements

```
pandas
openpyxl
datetime
warnings
os
```

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/wind-curtailment-analysis.git
cd wind-curtailment-analysis
```

2. Install required packages:
```bash
pip install pandas openpyxl
```

3. Create a `data` folder in the project directory

## Data Requirements

The tool expects the following Excel files in the `data/` folder:

### 1. Curtailment Data (`export-finished-YYYY-MM-DD HH_MM_SS.xlsx`)
- **Source**: German grid operators (TSOs)
- **Format**: Excel file with curtailment events
- **Required columns**:
  - `Start`: Start time (DD.MM.YYYY HH:MM:SS)
  - `Ende`: End time (DD.MM.YYYY HH:MM:SS)
  - `Dauer (Min)`: Duration in minutes
  - `Stufe (%)`: Curtailment level percentage
  - `Anlagenschlüssel`: Plant identifier

### 2. Market Price Data (`Gro_handelspreise_YYYYMMDDHHMMSS_YYYYMMDDHHMMSS_Stunde.xlsx`)
- **Source**: ENTSO-E or energy exchanges
- **Format**: Excel file with hourly electricity prices
- **Required columns**:
  - `Datum von`: Date/time (DD.MM.YYYY HH:MM)
  - `Deutschland/Luxemburg [€/MWh]`: Market price

### 3. Redispatch Data (`Ausgleichsenergie_YYYYMMDDHHMMSS_YYYYMMDDHHMMSS_Stunde.xlsx`)
- **Source**: German TSOs
- **Format**: Excel file with balancing energy prices
- **Required columns**:
  - `Datum von`: Date/time (DD.MM.YYYY HH:MM)
  - `Preis [€/MWh]`: Redispatch price

### 4. CO2 Emission Data (`stromix_YYYY.xlsx`)
- **Source**: Electricity Maps or similar providers
- **Format**: Excel file with hourly CO2 intensity
- **Required columns**:
  - `Datetime (UTC)`: UTC timestamp
  - `Carbon intensity direct (g/kWh)`: CO2 emission factor

## Configuration

Edit the configuration parameters at the top of the script:

```python
# Configuration Parameters
TURBINE_CAPACITY_MW = 2.3                              # Turbine capacity in MW
COMPENSATION_RATE = 0.925                              # Compensation rate (92.5%)
KWK1_ANLAGENSCHLUESSEL = 'E20793019000000000000004288000004'  # Plant identifier
```

Update file paths in the `FILES` dictionary:
```python
FILES = {
    'curtailment': 'your-curtailment-file.xlsx',
    'market': 'your-market-price-file.xlsx',
    'redispatch': 'your-redispatch-file.xlsx',
    'strommix': 'your-co2-data-file.xlsx'
}
```

## Usage

Run the analysis:

```bash
python curtailment_analysis.py
```

## Output

The tool generates a comprehensive report including:

### Main Results
- Total curtailed energy (MWh)
- Total missed revenue (€)
- Compensation paid (€)
- Total redispatch cost (€)
- Total economic impact (€)
- Total CO2 emissions (tonnes)
- Number of curtailment events
- Total duration of curtailments

### Key Insights
- Average curtailed energy per event
- Annual capacity factor loss
- Average electricity price during curtailment events
- Duration breakdown (minutes, hours, days)

### Example Output
```
=== CURTAILMENT IMPACT ANALYSIS RESULTS ===
Analysis Period: 2024
Wind Turbine: KWK 1 (Capacity: 2.3 MW)
Compensation Rate: 92.5%
--------------------------------------------------
Total Curtailed Energy (MWh): 8.36
Total Missed Revenue (€): 878.65
Compensation Paid (€): 812.75
Total Redispatch Cost (€): 622.15
Total Economic Impact (€): 1,434.90
Total CO2 Emissions (tonnes): 2.88
Number of Events: 112
Processed Events (with curtailment): 3
Total Duration (minutes): 1,200

=== KEY INSIGHTS ===
Average curtailed energy per event: 2.79 MWh
Annual capacity factor loss due to curtailment: 0.04%
Average electricity price during curtailment: 105.07 €/MWh
Total curtailment duration: 1,200 minutes (20.0 hours, 0.8 days)
```

## Methodology

### Curtailment Calculation
The tool calculates curtailed energy using:
```
Curtailed Energy = Duration × Turbine Capacity × Curtailment Percentage
```

### Economic Impact
- **Missed Revenue**: Curtailed energy × Market price
- **Compensation**: Missed revenue × Compensation rate (typically 92.5% in Germany)
- **Redispatch Cost**: Curtailed energy × Redispatch price

### Environmental Impact
- **CO2 Emissions**: Curtailed energy × CO2 emission factor
- Represents additional emissions from fossil fuel plants compensating for lost renewable energy

## Data Quality Features

- **Date Format Validation**: Handles German date formats (DD.MM.YYYY)
- **Missing Data Handling**: Graceful handling of missing values
- **Diagnostic Output**: Shows sample events for data validation
- **Duration Cross-Check**: Compares file duration with calculated duration

## Common Issues

### Large Duration Values
If you see unreasonably large total durations, check:
1. Events with 0% curtailment level (these are filtered out automatically)
2. Date parsing errors
3. Overlapping or incorrect time ranges in source data

### Missing Price Data
The tool handles missing market or redispatch price data by using zero values for affected periods.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- German TSOs for providing curtailment data
- ENTSO-E for electricity market data
- Electricity Maps for CO2 emission factors

## Contact

For questions or support, please open an issue on GitHub or contact [your-email@example.com].

## Version History

- **v1.0.0**: Initial release with basic curtailment analysis
- **v1.1.0**: Added CO2 emission calculation and improved data validation
- **v1.2.0**: Fixed duration calculation and added diagnostic features
