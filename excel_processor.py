#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel Column Extractor

This script reads data from an Excel file and creates a new Excel file
containing only the first three columns of the original file.
"""

import os
import sys
import pandas as pd
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger('excel_processor')

def process_excel_file(input_file_path, output_file_path=None):
    """
    Read an Excel file and create a new one with only the first three columns.
    
    Args:
        input_file_path (str): Path to the input Excel file
        output_file_path (str, optional): Path for the output Excel file.
            If not provided, will create in the same directory with '_processed' suffix.
    
    Returns:
        str: Path to the created output file
    """
    try:
        # Check if input file exists
        if not os.path.isfile(input_file_path):
            logger.error(f"Input file not found: {input_file_path}")
            return None
        
        # Check if file has a valid Excel extension
        valid_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.odf', '.ods', '.odt']
        _, file_ext = os.path.splitext(input_file_path)
        if file_ext.lower() not in valid_extensions:
            logger.error(f"Invalid file type: {file_ext}. Expected Excel file format.")
            return None
            
        # Generate output file path if not provided
        if output_file_path is None:
            file_name, file_ext = os.path.splitext(input_file_path)
            output_file_path = f"{file_name}_processed{file_ext}"
            
        logger.info(f"Reading Excel file: {input_file_path}")
        # Read the Excel file
        df = pd.read_excel(input_file_path)
        
        # Check if file has at least 3 columns
        if len(df.columns) < 3:
            logger.warning(f"Input file has fewer than 3 columns: {len(df.columns)} found")
            return None
            
        # Get the first three columns
        first_three_cols = df.iloc[:, :3]
        
        logger.info(f"Extracted first three columns: {first_three_cols.columns.tolist()}")
        
        # Write to a new Excel file
        first_three_cols.to_excel(output_file_path, index=False)
        
        logger.info(f"Successfully created output file: {output_file_path}")
        return output_file_path
        
    except Exception as e:
        logger.error(f"Error processing Excel file: {str(e)}", exc_info=True)
        return None

def main():
    """
    Main execution function. Handles command line arguments.
    """
    try:
        if len(sys.argv) < 2:
            logger.error("Missing required arguments")
            logger.info("Usage: python excel_processor.py <input_excel_file> [output_excel_file]")
            sys.exit(1)
            
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        
        result = process_excel_file(input_file, output_file)
        
        if result:
            logger.info(f"Successfully processed file. Output saved to: {result}")
            print(f"Successfully processed file. Output saved to: {result}")
            sys.exit(0)
        else:
            logger.error("Failed to process Excel file. Check the logs for details.")
            print("Failed to process Excel file. Check the logs for details.")
            sys.exit(1)
            
    except Exception as e:
        logger.error(f"Unexpected error in main execution: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()