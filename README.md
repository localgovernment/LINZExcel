Setup 
=====
1. Register with LINZ https://data.linz.govt.nz/
2. Obtain an API key http://www.linz.govt.nz/data/linz-data-service/guides-and-documentation/creating-an-api-key
3. Apply for restricted owner information datasets (this step may no longer be required)
4. Download LINZExcel.bas from https://raw.githubusercontent.com/localgovernment/LINZExcel/master/LINZExcel.bas
5. Create a new Excel spreadsheet 
6. Press ALT+F11
7. Press CTRL+M and import the LINZExcel.bas file
8. Open the LINZExcel module
9. Near the top of the code, find 'Public Const key as String = "my LINZ API key"' and replace key string with your LINZ api key
10. Close VBA for applications 
11. Creat a button in a worksheet and assign it to the GetTitleInformation macro.  See [here](http://office.microsoft.com/en-nz/excel-help/add-a-button-and-assign-a-macro-to-it-in-a-worksheet-HP010236676.aspx#BMadd_or_edit_a_button_(forms_toolbar)).

Using LINZExcel
===============
1. Enter valid titles into a column within the same worksheet as the button created during setup
2. Select the titles you want to process
3. Click the button
4. After a few seconds the following worksheets will be created and populated with data: PropertyTitlesList, PropertyTitleEstatesList, PropertyTitleOwnersList, TitleMemorialsList, TitleParcelAssociationList
5. The data received from LINZ is not ordered.  You'll need to use the Filter/Sort features of Excel to make sense of the results.
