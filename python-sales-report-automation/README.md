# Sales Report Automation With Python

A python project developed during my programming studies to analyze sales data from an excel spreadsheet and automatically send a sales report by email

## Features

- Reads sales data from an Excel spreadsheet
- Groups sales information by store
- Calculates total revenue from each store
- Calculates the total quantity of products sold
- Calculates the avarage revenue per product sold
- Generates HTML tables for the report
- Sends the report automatically through Microsoft Outlook

## Technologies Used

- Python
- Pandas
- OpenPyXL
- PyWin32
- Microsoft Excel
- Microsoft Outlook
- HTML

## How It Works

The program reads the 'vendas.xlsx' spreadsheet and processes the sales information using pandas.
It then calculates sales indicators for each store and converts the results into HTML tables
Finally, it creates an email containing the report and sends it through Microsoft Outlook

## Requirements

- Python installed
- Microsoft Outlook configured on the computer
- Requires Python libraries:

pip install pandas openpyxl pywin32

## How To Run

- Clone this repository
- Install the required dependencies
- place the 'vendas.xlsx' spreadsheet in the project directory
- Configure the recipient email address in the script
- Run the Python file.

## Learning Objectives

This project helped me practice:

- Data manipulation with Pandas
- Grouping and aggregating information
- Working with Excel spreadsheets
- Generating HTML reports
- Automating email communication with Python

## Credits

Especial thanks to 'Hashtag Treinamentos', who provided a free Python course who helped me at this project.

Developed as part of my Python programming studies