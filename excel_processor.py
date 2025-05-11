#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel Column Extractor

This script reads data from an Excel file and creates a new Excel file
containing only the first three columns of the original file.
"""

# Import necessary libraries
import os          # For file path operations
import sys         # For command-line arguments and exit codes
import pandas as pd  # For Excel file manipulation
import logging     # For application logging

# Set up logging configuration
logging.basicConfig(
    level=logging.INFO,  # Set minimum log level to INFO
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',  # Define log format with timestamp
    handlers=[
        logging.StreamHandler(sys.stdout)  # Output logs to standard output
    ]
)
# Create a logger instance for this module
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
        
        # Define valid Excel file extensions
        valid_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.odf', '.ods', '.odt']
        # Split the file path to get the extension
        _, file_ext = os.path.splitext(input_file_path)
        # Validate if the file has a proper Excel extension
        if file_ext.lower() not in valid_extensions:
            logger.error(f"Invalid file type: {file_ext}. Expected Excel file format.")
            return None
            
        # Generate output file path if not provided by appending '_processed' to the filename
        if output_file_path is None:
            file_name, file_ext = os.path.splitext(input_file_path)
            output_file_path = f"{file_name}_processed{file_ext}"
            
        logger.info(f"Reading Excel file: {input_file_path}")
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(input_file_path)
        
        # Ensure the file has at least 3 columns to extract
        if len(df.columns) < 3:
            logger.warning(f"Input file has fewer than 3 columns: {len(df.columns)} found")
            return None
            
        # Extract only the first three columns using iloc to select by position
        first_three_cols = df.iloc[:, :3]
        
        # Log the column names that were extracted
        logger.info(f"Extracted first three columns: {first_three_cols.columns.tolist()}")
        
        # Save the extracted columns to a new Excel file
        first_three_cols.to_excel(output_file_path, index=False)
        
        logger.info(f"Successfully created output file: {output_file_path}")
        return output_file_path
        
    except Exception as e:
        # Catch and log any exceptions that occur during processing
        logger.error(f"Error processing Excel file: {str(e)}", exc_info=True)
        return None

def main():
    """
    Main execution function. Handles command line arguments.
    """
    try:
        # Check if required command line arguments are provided
        if len(sys.argv) < 2:
            logger.error("Missing required arguments")
            logger.info("Usage: python excel_processor.py <input_excel_file> [output_excel_file]")
            sys.exit(1)  # Exit with error code
            
        # Get input file path from first argument
        input_file = sys.argv[1]
        # Get output file path from second argument if provided
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        
        # Process the Excel file
        result = process_excel_file(input_file, output_file)
        
        # Handle success or failure of the processing
        if result:
            logger.info(f"Successfully processed file. Output saved to: {result}")
            print(f"Successfully processed file. Output saved to: {result}")  # User-friendly output
            sys.exit(0)  # Exit with success code
        else:
            logger.error("Failed to process Excel file. Check the logs for details.")
            print("Failed to process Excel file. Check the logs for details.")  # User-friendly error
            sys.exit(1)  # Exit with error code
            
    except Exception as e:
        # Catch and log any unexpected exceptions in the main function
        logger.error(f"Unexpected error in main execution: {str(e)}", exc_info=True)
        sys.exit(1)  # Exit with error code

# Execute main function if script is run directly (not imported)
if __name__ == "__main__":
    main()