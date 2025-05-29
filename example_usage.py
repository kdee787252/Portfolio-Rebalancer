# -*- coding: utf-8 -*-
"""
Created on Fri May 30 00:04:50 2025

@author: arush
"""

from index_rebalancer import IndexRebalance

def main():
    # Initialize rebalancer
    rebalancer = IndexRebalance()
    
    # Load data for specific date
    rebalancer.load_data('securities_sample.xlsx', '06/30/17')
    
    # Rebalance portfolio
    portfolio = rebalancer.rebalance()
    
    # Display results
    print("\nRebalanced Portfolio:")
    print(portfolio[['Name', 'RIC', 'Weight']])
    
    # Export to Excel
    output_file = rebalancer.export_results()
    print(f"\nResults exported to: {output_file}")

if __name__ == "__main__":
    main()