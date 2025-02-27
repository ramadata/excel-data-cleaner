import os
import logging
import pandas as pd
import numpy as np
from datetime import datetime



def setup_logger():
    """Set up and configure logger"""
    # Create logger
    logger = logging.getLogger('data_quality')
    logger.setLevel(logging.INFO)
    
    # Create console handler and set level
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Create file handler and set level
    file_handler = logging.FileHandler('data_quality_{date_now}.log'.format(date_now=datetime.now()))
    file_handler.setLevel(logging.INFO)
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # Add formatter to handlers
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    
    # Add handlers to logger
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    
    return logger

def improve_excel_data_quality(file_path, output_path=None, log_level=logging.INFO):
    """
    Process an Excel file to improve data quality by handling common issues
    
    Parameters:
    -----------
    file_path : str
        Path to the input Excel file
    output_path : str, optional
        Path to save the cleaned Excel file. If None, will add '_cleaned' to the original filename
    log_level : logging level, optional
        The level of logging detail
    
    Returns:
    --------
    pd.DataFrame
        The cleaned dataframe
    """
    # Setup logger
    logger = setup_logger()
    logger.setLevel(log_level)
    
    # Log start of processing
    logger.info("Starting data quality improvement process")
    
    # Determine output path if not provided
    if output_path is None:
        file_name, file_ext = os.path.splitext(file_path)
        output_path = f"{file_name}_cleaned{file_ext}"
    
    logger.info(f"Loading file: {file_path}")
    
    # Read the Excel file
    try:
        df = pd.read_excel(file_path)
        logger.info(f"Successfully loaded file with {len(df)} rows and {len(df.columns)} columns")
    except Exception as e:
        logger.error(f"Error reading file: {e}")
        return None
    
    # Store original row count
    original_row_count = len(df)
    logger.info(f"Original dataset: {original_row_count} rows, {len(df.columns)} columns")
    
    # Handle column names - strip whitespace, replace spaces with underscores, lowercase
    old_columns = list(df.columns)
    df.columns = [col.strip().lower().replace(' ', '_') for col in df.columns]
    logger.info("Standardized column names")
    logger.debug(f"Column names before: {old_columns}")
    logger.debug(f"Column names after: {list(df.columns)}")
    
    # Remove duplicate rows
    logger.info("Checking for duplicate rows")
    df_deduped = df.drop_duplicates()
    dupes_removed = original_row_count - len(df_deduped)
    logger.info(f"Removed {dupes_removed} duplicate rows")
    
    # Handle missing values appropriately for each column
    logger.info("Handling missing values...")
    for column in df_deduped.columns:
        missing_count = df_deduped[column].isna().sum()
        if missing_count > 0:
            logger.info(f"Column '{column}' has {missing_count} missing values")
            
            # Numeric columns: fill with median
            if pd.api.types.is_numeric_dtype(df_deduped[column]):
                median_val = df_deduped[column].median()
                df_deduped[column].fillna(median_val, inplace=True)
                logger.info(f"Filled column '{column}' with median: {median_val}")
            
            # Date columns: identified by name or type
            elif ('date' in column.lower() or 
                  pd.api.types.is_datetime64_dtype(df_deduped[column])):
                # Fill with the most recent date before the missing value
                df_deduped[column] = df_deduped[column].ffill()
                logger.info(f"Filled dates in column '{column}' using forward fill")
            
            # Categorical/text: fill with mode
            else:
                mode_val = df_deduped[column].mode()[0] if not df_deduped[column].mode().empty else "Unknown"
                df_deduped[column].fillna(mode_val, inplace=True)
                logger.info(f"Filled column '{column}' with mode: {mode_val}")
    
    # Handle outliers in numeric columns
    logger.info("Detecting and handling outliers...")
    for column in df_deduped.select_dtypes(include=[np.number]).columns:
        # Calculate IQR
        Q1 = df_deduped[column].quantile(0.25)
        Q3 = df_deduped[column].quantile(0.75)
        IQR = Q3 - Q1
        
        # Define outlier boundaries
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        
        # Count outliers
        outliers = df_deduped[(df_deduped[column] < lower_bound) | 
                              (df_deduped[column] > upper_bound)]
        
        if len(outliers) > 0:
            logger.info(f"Column '{column}' has {len(outliers)} outliers")
            
            # Cap outliers at the boundaries
            df_deduped[column] = df_deduped[column].clip(lower=lower_bound, upper=upper_bound)
            logger.info(f"Capped outliers in '{column}' between {lower_bound:.2f} and {upper_bound:.2f}")
    
    # Standardize date formats
    logger.info("Standardizing date formats...")
    for column in df_deduped.columns:
        # Check if column might contain dates based on name
        if any(date_hint in column.lower() for date_hint in ['date', 'time', 'day', 'month', 'year']):
            try:
                # Try to convert to datetime
                df_deduped[column] = pd.to_datetime(df_deduped[column], errors='coerce')
                logger.info(f"Converted '{column}' to datetime format")
            except Exception as e:
                logger.warning(f"Could not convert '{column}' to datetime: {e}")
    
    # Standardize text case for string columns
    logger.info("Standardizing text formats...")
    for column in df_deduped.select_dtypes(include=['object']).columns:
        # Check if column looks like it contains names or titles
        if any(name_hint in column.lower() for name_hint in ['name', 'title', 'category', 'type']):
            # Title case for names and categories
            df_deduped[column] = df_deduped[column].astype(str).str.title()
            logger.info(f"Converted '{column}' to title case")
        else:
            # Otherwise, lowercase for consistency
            df_deduped[column] = df_deduped[column].astype(str).str.strip().str.lower()
            logger.info(f"Standardized '{column}' to lowercase")
    
    # Add data quality metrics columns
    df_deduped['row_completeness'] = df_deduped.notna().mean(axis=1)
    logger.info("Added row completeness score column")
    
    # Validate email addresses if any columns might contain them
    for column in df_deduped.columns:
        if 'email' in column.lower():
            logger.info(f"Validating email addresses in '{column}'...")
            # Simple regex for email validation
            email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            
            # Create a validation column
            validation_col = f"{column}_valid"
            df_deduped[validation_col] = df_deduped[column].astype(str).str.match(email_pattern)
            
            invalid_count = (~df_deduped[validation_col]).sum()
            logger.info(f"Found {invalid_count} invalid email addresses in '{column}'")
    
    # Save the cleaned data
    try:
        df_deduped.to_excel(output_path, index=False)
        logger.info(f"Cleaned data saved to {output_path}")
    except Exception as e:
        logger.error(f"Error saving cleaned data: {e}")
    
    # Generate a data quality report
    quality_report = {
        'original_rows': original_row_count,
        'cleaned_rows': len(df_deduped),
        'duplicates_removed': dupes_removed,
        'columns_processed': len(df_deduped.columns),
        'overall_completeness': df_deduped.notna().mean().mean() * 100
    }
    
    logger.info("Data Quality Report:")
    for metric, value in quality_report.items():
        if 'completeness' in metric:
            logger.info(f"{metric.replace('_', ' ').title()}: {value:.2f}%")
        else:
            logger.info(f"{metric.replace('_', ' ').title()}: {value}")
    
    logger.info("Data quality improvement process completed")
    return df_deduped

# Example usage
if __name__ == "__main__":
    # Replace with your file path
    input_file = "your_data.xlsx"
    
    try:
        # Set to logging.DEBUG for more detailed logs
        cleaned_df = improve_excel_data_quality(input_file, log_level=logging.INFO)
        
        # Additional analysis on the cleaned data
        if cleaned_df is not None:
            logger = logging.getLogger('data_quality')
            logger.info("Generating column-wise completeness metrics")
            completeness = cleaned_df.notna().mean() * 100
            for col, comp in completeness.items():
                logger.info(f"Column completeness - {col}: {comp:.2f}%")
    except FileNotFoundError:
        logger = logging.getLogger('data_quality')
        logger.error(f"File not found: {input_file}")
        logger.error("Please update the script with the correct file path.")