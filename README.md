# FINRA Data Pull and Excel Report Generator

This script pulls short interest data from the FINRA API and generates an Excel table with specific filters and columns as specified. 

## Features

- Accepts a date input from the user.
- Filters data based on average daily volume, exchange code, and accounting year month.
- Generates an Excel table with specified columns.

## Script Details

### Overview

The script performs the following key steps:

1. **User Input:** Asks the user to input a date each time the script runs.
2. **API Request:** Makes a request to the FINRA API with specified filters.
3. **Data Processing:** Processes the received data and extracts necessary columns.
4. **Excel Generation:** Creates an Excel table with the processed data.

## Notes

- Ensure you have an active internet connection to pull data from the FINRA API.
- The script assumes the FINRA API endpoint and the format of the data to be consistent. Any changes in the API might require modifications to the script.
