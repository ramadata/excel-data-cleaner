# Excel Data Quality Improvement Tool

A Python-based utility for automatically improving data quality in Excel files through cleaning, standardization, and validation.

## Overview

This tool performs automated data quality improvements on Excel files, handling common data issues such as:

- Missing values
- Duplicate records
- Inconsistent formatting
- Outliers
- Invalid data entries
- Non-standardized text and date formats

All operations are thoroughly logged for traceability and auditability.

## Features

- **Column name standardization**: Converts column names to lowercase with underscores
- **Duplicate detection and removal**: Identifies and eliminates duplicate records
- **Intelligent missing value handling**:
  - Fills numeric fields with median values
  - Forward-fills date fields
  - Fills categorical/text fields with mode values
- **Outlier detection and handling** using the IQR method
- **Date format standardization**
- **Text standardization**:
  - Title case for names, categories, and titles
  - Lowercase for other text fields
- **Data validation**:
  - Email format validation
  - Row completeness scoring
- **Comprehensive logging** of all operations and changes
- **Data quality reporting** with key metrics

## Requirements

- Python 3.6+
- pandas
- numpy


## Usage

### Basic Usage

```python
from excel_data_cleaner import improve_excel_data_quality

# Process a file with default settings
improve_excel_data_quality("path/to/your/data.xlsx")
```

This will:
1. Process the file at the specified path
2. Create a new file with "_cleaned" appended to the original filename
3. Generate a log file with details of all operations

### Advanced Usage

```python
import logging
from excel_data_cleaner import improve_excel_data_quality

# Specify output path and logging level
cleaned_df = improve_excel_data_quality(
    file_path="path/to/your/data.xlsx",
    output_path="path/to/output/cleaned_data.xlsx",
    log_level=logging.DEBUG  # For detailed logs
)

# Perform additional operations on the cleaned DataFrame
if cleaned_df is not None:
    # Your custom operations here
    pass
```

### Logging

The script generates detailed logs in both the console and a `data_quality.log` file. You can adjust the verbosity with the `log_level` parameter:

- `logging.DEBUG`: Most detailed, includes all operations
- `logging.INFO`: Standard information (default)
- `logging.WARNING`: Only potential issues
- `logging.ERROR`: Only errors

## Customization

You can extend the script by:

1. Adding custom validation rules
2. Implementing domain-specific cleaning functions
3. Modifying the outlier detection thresholds
4. Adding new data quality metrics

## Data Quality Report

After processing, the tool generates a data quality report with metrics including:

- Original row count
- Cleaned row count
- Number of duplicates removed
- Number of columns processed
- Overall data completeness percentage
- Column-wise completeness percentages

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
