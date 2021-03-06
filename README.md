Setup 
=====
1. Register with LINZ https://data.linz.govt.nz/
2. Obtain an API key http://www.linz.govt.nz/data/linz-data-service/guides-and-documentation/creating-an-api-key
3. Apply for restricted owner information datasets
4. Download LINZExcel.bas from https://raw.githubusercontent.com/localgovernment/LINZExcel/master/LINZExcel.bas
5. Create a new Excel spreadsheet 
6. Press ALT+F11
7. Press CTRL+M and import the LINZExcel.bas file
8. Open the LINZExcel module
9. Near the top of the code, find 'Public Const key as String = "my LINZ API key"' and replace key string with your LINZ api key
10. Close VBA for applications 
11. Creat a button in a worksheet and assign it to the GetTitleInformation macro.  See [here](https://support.office.com/en-us/article/Add-a-button-and-assign-a-macro-to-it-in-a-worksheet-d58edd7d-cb04-4964-bead-9c72c843a283?CorrelationId=d44b2204-cdf2-4e1a-98e0-9dfed6cb47f7&ui=en-US&rs=en-US&ad=US&ocmsassetID=HP010236676#bmadd_or_edit_a_button__forms_toolbar_).
12. Save as an 'Excel Macro-Enabled Workbook' (.xlsm)

Using LINZExcel
===============
1. Enter valid titles into a column within the same worksheet as the button created during setup
2. Select the titles you want to process
3. Click the button
4. After a few seconds the following worksheets will be created and populated with data: PropertyTitlesList, PropertyTitleEstatesList, PropertyTitleOwnersList, TitleMemorialsList, TitleParcelAssociationList
5. The data received from LINZ is not ordered.  You'll need to use the Filter/Sort features of Excel to make sense of the results.
