# Portfolio Rebalancer

A Python tool for rebalancing investment portfolios based on market capitalization.

## Features
- Loads security data from Excel files
- Filters securities by target rebalance date
- Calculates MCAP-based weights
- Exports results to Excel with auto-generated filenames

## Usage
1. Install dependencies: `pip install -r requirements.txt`
2. Prepare Excel file with columns: Date, Name, RIC, MCAP
3. Run example: `python example_usage.py`

```python
from index_rebalancer import IndexRebalance

rebalancer = IndexRebalance()
rebalancer.load_data('securities.xlsx', '06/30/17')
portfolio = rebalancer.rebalance()
rebalancer.export_results()