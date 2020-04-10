# Bulk Deals Stock Exchanges (BSE & NSE)

## Description

Financial data output for all large volume deals executed for any data and Indian Stock Exchange (BSE/NSE).
***Utility to fetch data for bulk deals for BSE and NSE***

## Steps for execution

Run the following command in the command prompt
 BulkDealsSensex.exe start date end date BSE | NSE | both>

## Condition/Restrictions

 Start Date / End Date supported FORMAT = DD-MM-YYYY
 ***No other format will work.***

Today's date is not supported in any condition as data is only published at the source only the next calendar day.

## Result

 File: BSE.xlsx || NSE.xlsx || BSE_NSE_Bulk Deals.xlsx will will be created in the "Documents" folder.

## Sample Commands

 BulkDealsSensex.exe 01-04-2020 10-04-2020 BSE

 BulkDealsSensex.exe 01-04-2020 10-04-2020 NSE

 BulkDealsSensex.exe 01-04-2020 10-04-2020 both

## Change Log - Version 2.0

- Added support for dates with formating inside the output excel files.

- Updated logic to output only a single file when calling the program with "both" argument.
  - The data for respective exchanges will be represented in different tabs.

- Upgraded all packages in the program
  - The program now is compatible with Chrome 81 and Chromedriver version can be downloaded from here: 
  <https://chromedriver.storage.googleapis.com/index.html?path=81.0.4044.69/>

- Added a copy of the application into the codebase as well.
  - You can download the release version 2.0 zip file from the 'Release Build' folder

## Known Issues

One known issue is for the presence of the error message for timeout in the application dialog. This is thrown from the chromedriver.exe.
Application is still running perfectly in the background.
