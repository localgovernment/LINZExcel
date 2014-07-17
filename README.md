LINZExcel
=========
A collection of Excel VBA functions that access the LINZ data.linz.govt.nz WFS API

Setup 
=====
1. Register with LINZ https://data.linz.govt.nz/
2. Obtain an API key http://www.linz.govt.nz/about-linz/linz-data-service/features/how-to-use-web-services
3. Apply for restricted datasets https://data.linz.govt.nz/login/?next=/group/restricted-owner-data-group/request-access/
4. Download LINZExcel.bas from https://raw.githubusercontent.com/localgovernment/LINZExcel/master/LINZExcel.bas
5. Create a new Excel spreadsheet 
6. Press ALT+F11
7. Press CTRL+M and import the LINZExcel.bas file
8. Open the LINZExcel module
9. Near the top of the code, find 'Public Const key as String = "my LINZ API key"' and replace key string with your LINZ api key
10. Close VBA for applications 

Example
=======
As an example populate the following cells in the worksheet as follows:

* A1: Parcel ID
* B1: Title(s)
* C1: Encumbrancees
* D1: Instrument Numbers
* E1: Instrument Types
* B2: =GetTitles(A2)
* C2: =GetEncumbrancees(B2)
* D2: =GetInstrumentNumbers(B2)
* E2: =GetInstrumentTypes(B2)

Enter a valid Parcel ID into A2 and the rest of the cells in row 2 should populate automatically.

To save: File 'Save As' then Save as Type 'Excel Macro-enabled Workbook' (xlsm)

Notes
=====
* GetTitles(parcelID) - returns a list of (current) LINZ titles for the given ParcelID
* GetEncumbrancees(title) - returns a list of (current) LINZ encumbrancees for the given title
* GetInstrumentNumbers(title) - returns a list of (current) LINZ Instrument numbers for the given title
* GetInstrumentTypes(title) - returns a list of (current) LINZ instrument types for the given title – should be same order as instrument numbers
* GetMainParcelID(valuation) - Taupo District Council Mapi API - returns the main parcel ID for the given valuation number
* GetSurnames - returns a list of LINZ surnames for the given title
* ProcessValuations - Process all selected valuations - TDC specific but can be modified.  Assign to a button.

GetEncumbrancees, GetInstrumentNumbers, and GetIntrumentTypes will take the first title in the list (if a list of titles  is provided)
 
