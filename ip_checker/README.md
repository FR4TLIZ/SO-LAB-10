# Overview
This PowerShell script checks the availability of the first 5 IP addresses listed in an Excel file. 
For each address, it performs a network check (ping) and records whether the address is accessible or not. 
The results are saved to a new Excel file.

# Prerequisites
- The ImportExcel module for PowerShell (used for reading and writing Excel files)
- Command: Install-Module -Name ImportExcel -Force
- An Excel file (adresy.xlsx) containing a list of IP addresses.

# Running the Script
Make sure you have adresy.xlsx with the IP addresses.
Run the script in PowerShell: .\LAB_10.ps1
The script will:
- Check the availability of the first 5 IP addresses.
- Output the results to wyniki.xlsx.
