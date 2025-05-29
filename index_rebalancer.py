# -*- coding: utf-8 -*-
"""
Created on Fri May 30 00:00:57 2025

@author: arush
"""

import pandas as pd

class IndexRebalance:
    """
    A portfolio rebalancing tool that calculates market capitalization-based weights
    
    Attributes:
        data (DataFrame): Raw security data loaded from Excel
        constituents (DataFrame): Securities filtered for target date
        rebalanced_df (DataFrame): Final portfolio with calculated weights
    """
    
    def __init__(self):
        """Initialize a new rebalancer instance"""
        self.data = None
        self.constituents = None
        self.rebalanced_df = None
    
    def load_data(self, input_file, rebalance_date):
        """
        Load security data and filter for specific rebalance date
        
        Args:
            input_file (str): Path to Excel file with security data
            rebalance_date (str): Target date for rebalancing (format: MM/DD/YY)
            
        Returns:
            DataFrame: Filtered securities for target date
        """
        self.data = pd.read_excel(input_file)
        self.constituents = self.data[self.data['Date'] == rebalance_date]
        return self.constituents
    
    def rebalance(self):
        """
        Calculate MCAP-based weights for the portfolio
        
        Returns:
            DataFrame: Portfolio with Date, Name, RIC, MCAP, and Weight columns
        """
        total_mcap = self.constituents['MCAP'].sum()
        self.rebalanced_df = self.constituents.copy()
        self.rebalanced_df['Weight'] = (self.rebalanced_df['MCAP'] / total_mcap) * 100
        return self.rebalanced_df
    
    def export_results(self, output_file=None):
        """
        Export rebalanced portfolio to Excel
        
        Args:
            output_file (str, optional): Custom output filename. Defaults to auto-generated name.
            
        Returns:
            str: Path to the generated Excel file
        """
        if output_file is None:
            date_str = self.rebalanced_df['Date'].iloc[0]
            output_file = f"rebalanced_portfolio_{date_str.replace('/', '-')}.xlsx"
        
        self.rebalanced_df.to_excel(output_file, index=False)
        return output_file