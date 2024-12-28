# README: BusyWin Tools

## Overview

`busywin_tools.py` is a Python script designed to process and extract transactional data from BusyWin `.bds` database files. It provides functionality to connect to BusyWin databases, retrieve transactional data, process it into a unified format, and optionally export the processed data to a CSV file.

## Features

- Automatically detects `.bds` database files in a specified folder.
- Extracts transactional data from BusyWin databases using predefined SQL queries.
- Processes the data to generate a comprehensive DataFrame containing detailed transaction information.
- Exports the processed data as a CSV file.

## Installation

1. Ensure you have the following prerequisites installed:
   - Python 3.6+
   - Required Python packages: `pandas`, `pyodbc`
2. Install the required packages:
   ```bash
   pip install pandas pyodbc
   ```

## Usage

### 1. Setup the Folder

- Place all `.bds` files in a folder named `busywin_data` or specify a custom folder during initialization.

### 2. Running the Script

Create a Python script (`main.py`) to use the `busywin_transactions` class:

```python
from busywin_tools import busywin_transactions

data = busywin_transactions()
df = data.get_data()
print(df)

# To export the transaction data to a CSV file
# data.export_data()
```

### 3. Output

The output will be a DataFrame with the following columns:

| Column Name       | Description                              |
|--------------------|------------------------------------------|
| `vchtype`         | Voucher type (e.g., purchase, sales)     |
| `vchname`         | Voucher name                            |
| `vchcode`         | Voucher code                            |
| `vchno`           | Voucher number                          |
| `party`           | Party involved in the transaction       |
| `rectype`         | Record type                             |
| `sno`             | Serial number                           |
| `date`            | Date and time of the transaction        |
| `code`            | Item code                               |
| `codetype`        | Type of code                            |
| `value`           | Value of the transaction                |
| `name`            | Name of the item/party                 |
| `parentgrp`       | Parent group                            |
| `hsn`             | HSN code                                |
| `qty`             | Quantity                                |
| `unit`            | Unit of measurement                    |
| `mrp`             | Maximum retail price                   |
| `listprice`       | List price                              |
| `discountper`     | Discount percentage                     |
| `price`           | Price after discount                   |
| `amount`          | Amount before taxes                    |
| `netamount`       | Net amount after taxes                 |
| `cgstper`         | Central GST percentage                 |
| `cgst`            | Central GST amount                     |
| `sgstper`         | State GST percentage                   |
| `sgst`            | State GST amount                       |
| `cessper`         | CESS percentage                        |
| `cess`            | CESS amount                            |
| `taxperwocess`    | Tax percentage without CESS            |
| `taxwocess`       | Tax amount without CESS                |
| `taxper`          | Total tax percentage                   |
| `tax`             | Total tax amount                       |

### 4. Exporting Data

To export the processed data, use:
```python
data.export_data()
```

This will create a CSV file named `busywin_transactions.csv` in the working directory.

## Configuration

You can customize the following parameters during initialization:

- **`busywin_folder`**: The folder where `.bds` files are located (default: `busywin_data`).
- **`export_file_name`**: The name of the exported CSV file (default: `busywin_transactions`).

Example:
```python
data = busywin_transactions(busywin_folder="custom_folder", export_file_name="custom_file_name")
```

## Notes

1. Ensure the Microsoft Access ODBC driver is installed on your system to enable database connections.
2. Move all `.bds` files to the specified folder before running the script.

## License

This project is licensed under the MIT License. 

For any issues or contributions, feel free to create a pull request or open an issue in the repository.
