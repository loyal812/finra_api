import requests
import pandas as pd
from datetime import datetime
import os
from dotenv import load_dotenv

# Load env data from .env file
load_dotenv()

def retrieve_api_token(finra_api_key, finra_api_secret):
    """
    Load credentials to access FINRA API token.
    """
        
    my_token = requests.post('https://ews.fip.finra.org/fip/rest/ews/oauth2/access_token?grant_type=client_credentials',
                            auth = (finra_api_key, finra_api_secret))

    assert my_token.status_code == 200, "Access Token could not be generated"

    global my_access_token
    my_access_token = my_token.json()['access_token']

    return my_access_token

def get_finra_data(date_input, token):
    # Define the base URL for the FINRA API
    base_url = "https://api.finra.org/data/group/otcMarket/name/consolidatedShortInterest"
    
    # Define the parameters
    params = {
        "limit": 5000,
        "compareFilters": [
          {
              "compareType": "EQUAL",
              "fieldName": "accountingYearMonthNumber",
              "fieldValue": date_input
          },
          {
              "compareType": "GTE",
              "fieldName": "averageDailyVolumeQuantity",
              "fieldValue": "500000"
          }
        ],
        "orFilters": [
            {
                "compareFilters": [
                    {
                        "compareType": "EQUAL",
                        "fieldName": "marketClassCode",
                        "fieldValue": "NYSE"
                    },
                    {
                        "compareType": "EQUAL",
                        "fieldName": "marketClassCode",
                        "fieldValue": "NNM"
                    },
                ]
            }
        ]
    }
    
    # Define the headers including the API key
    headers = {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }
    
    # Perform the API request
    response = requests.post(base_url, headers=headers, json=params)
    
    # Check if the request was successful
    if response.status_code != 200:
        raise Exception(f"Failed to retrieve data: {response.status_code} - {response.text}")
    
    # Parse the JSON response
    data = response.json()
    
    # Extract the relevant data
    extracted_data = []
    
    for record in data:
        extracted_data.append({
            "Symbol": record.get("symbolCode"),
            "Company Name": record.get("issueName"),
            "Days to Cover": record.get("daysToCoverQuantity"),
            "marketClassCode": record.get("marketClassCode"),
            "accountingYearMonthNumber": record.get("accountingYearMonthNumber")
        })
    
    return extracted_data

def save_to_excel(data, date_input):
    # Create a DataFrame
    df = pd.DataFrame(data)
    
    # Define the Excel file name
    file_name = f"FINRA_Data_{date_input}.xlsx"
    
    # Save the DataFrame to an Excel file
    df.to_excel(file_name, index=False)
    
    print(f"Data successfully saved to {file_name}")

def main():
    # Ask the user to input the date
    date_input = input("Please enter the date (e.g., 20240614): ")

    finra_api_key = os.getenv("FINRA_API_KEY")
    finra_api_secret = os.getenv("FINRA_API_SECRET")

    my_access_token = retrieve_api_token(finra_api_key, finra_api_secret)

    try:
        # Validate the date input
        datetime.strptime(date_input, '%Y%m%d')
    except ValueError:
        print("Invalid date format. Please use YYYYMMDD.")
        return
    
    # Get the data from the FINRA API
    data = get_finra_data(date_input, my_access_token)
    
    # Save the data to an Excel file
    save_to_excel(data, date_input)

if __name__ == "__main__":
    main()
