# Bulk Deals Stock Exchanges (BSE & NSE)

## Description

Financial data output for all large volume deals executed for any data and Indian Stock Exchanges (**BSE/NSE**).

**_This is a utility application_**

## Technology Stack

| Tech              |   Name with Version   |
| ----------------- | :-------------------: |
| Language          |          C#           |
| Framework         |       .NET v4.5       |
| Automation        | Selenium WebDriver v3 |
| Excel Interaction |       EPPlus v5       |
| IDE               |  Visual Studio 2019   |
| Platform          |        Windows        |

## Steps for execution

1. Open a Command prompt window.
2. Navigate to the folder containing the application.
3. Run the following command in the command prompt.

`BulkDealsSensex.exe <start-date> <end date> <BSE || NSE || both>`

## Condition/Restrictions

- Start Date || End Date is only supported in **DD-MM-YYYY** format.

- Today's date is not supported:
  - Data for today's date is published at the source on the next calendar day.

## Result

Post completion of the operation successfully, an XLSX file will be created inside the **Documents** folder.

`File: BSE.xlsx || NSE.xlsx || BSE_NSE_Bulk_Deals.xlsx`

## Sample Commands

`BulkDealsSensex.exe 01-04-2020 10-04-2020 BSE`

`BulkDealsSensex.exe 01-04-2020 10-04-2020 NSE`

`BulkDealsSensex.exe 01-04-2020 10-04-2020 both`

## Change Log - Version 2.0

- Added support to run the application standalone without providing commandline arguments.

  - When the application is executed without giving the arguments, the data between `Yesterday` and `Day before Yesterday` will be extracted.

- Added support for dates with formating inside the output excel files.

- Updated logic to output only a single file when calling the program with "both" argument.

  - Respective exchanges data will be created in different tabs.

- Upgraded all packages in the program to latest versions.

  - Selenium Webdriver
  - EPPlus

- Compatible with the latest version of Chrome (Version 81.0.4044.92 (Official Build) (64-bit) at the time of publishing)

  - Download supported chromedriver: [Chromedriver version 81.0.4044.69](https://chromedriver.storage.googleapis.com/index.html?path=81.0.4044.69/) from here.

- Added a copy of the application into the codebase as well.
  - You can download the release version 2.0 zip file from the 'Release Build' folder

## Known Issue

One known issue is for the presence of the error message for timeout in the application dialog. This is thrown from the `chromedriver.exe`.

Application is still running perfectly in the background.

[More info for the issue can be found on stackoverflow.](https://stackoverflow.com/questions/60114639/timed-out-receiving-message-from-renderer-0-100-log-messages-using-chromedriver)
