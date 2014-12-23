' ************************************************************************************************
' * LINZExcel : Extract title information and related records from LINZ
' *
' * The MIT License (MIT)
' *
' * Copyright (c) 2014 Taupo District Council, Dion Liddell.
' ************************************************************************************************

' ******* Replace the following with your LINZ API Key *********
Public Const key As String = "my LINZ API KEY"
Public Sub GetTitleInformation()
    ' Generate all sheets for selected titles
    Dim selected As Range
    Set selected = Selection
    
    GetPropertyTitlesList selected
    GetPropertyTitleEstatesList selected
    GetPropertyTitleOwnersList selected
    GetTitleMemorialsList selected
    GetTitleParcelAssociationList selected
End Sub
Public Sub GetPropertyTitlesList(selected)
    ' This table provides live and part cancelled Title information
    Dim csvString As String: csvString = LINZCSV("table-1567", LINZCSVFilter("title_no", selected, ""))
    CSVtoSheet "PropertyTitlesList", csvString
    
End Sub
Public Sub GetPropertyTitleEstatesList(selected)
    ' A title estate is a type of ownership of a piece of land e.g. fee simple estate, leasehold estate.
    ' Estates are used to link the owners to the title. A title can have more than 1 estate and type.
    Dim csvString As String: csvString = LINZCSV("table-1566", LINZCSVFilter("title_no", selected, ""))
    CSVtoSheet "PropertyTitleEstatesList", csvString
End Sub
Public Sub GetPropertyTitleOwnersList(selected)
    ' This table provides registered (or current) ownership information for a Title. An owner (or proprietor) is a person
    ' or corporation holding a share in a Title estate.
    Dim csvString As String: csvString = LINZCSV("table-1564", LINZCSVFilter("title_no", selected, ""))
    CSVtoSheet "PropertyTitleOwnersList", csvString
End Sub
Public Sub GetTitleMemorialsList(selected)
    ' A title memorial is information recorded on a property title relating to a transaction, interest or restriction over a
    ' piece of land. Memorials can include details of mortgages, discharge of mortgages, transfer of ownership, and
    ' leases; all of which affect the land in some way.
    Dim csvString As String: csvString = LINZCSV("table-1695", LINZCSVFilter("title_no", selected, "current=%27true%27"))
    CSVtoSheet "TitleMemorialsList", csvString
End Sub
Public Sub GetTitleParcelAssociationList(selected)
    ' This table is used to associate live and part cancelled titles to current spatial parcels. There is a many to many relationship between titles and parcels
    Dim csvString As String: csvString = LINZCSV("table-1569", LINZCSVFilter("title_no", selected, ""))
    CSVtoSheet "TitleParcelAssociationList", csvString
End Sub
Public Sub CSVtoSheet(worksheetName, csvString)
    ' Convert CSV string to worksheet
    
    Dim destSheet As Worksheet
    Set destSheet = Sheets.Add(After:=Worksheets(Worksheets.Count))
    destSheet.Name = worksheetName & Worksheets.Count
    
    csvLines = Split(csvString, Chr(10))
    For csvLine = 0 To UBound(csvLines)
        ' csvFields = Split(csvLines(csvLine), ",")
        csvFields = CleanSplit(csvLines(csvLine))
        For csvField = 0 To UBound(csvFields)
            destSheet.Range("A1").Offset(csvLine, csvField).Value = csvFields(csvField)
        Next csvField
    Next csvLine
    
End Sub
Public Function CleanSplit(csvLine)
    ' Split a CSV line into fields by comma delimiter excluding those commas within quotes
    Dim csvFieldIndex As Integer
    Dim cleanedCSVFields() As String: ReDim cleanedCSVFields(0)
    Dim cleanedCSVFieldsIndex As Integer
    Dim cleanedField As String
    Dim concatFieldIndex As Integer
    
    csvFields = Split(csvLine, ",")
    
    csvFieldIndex = 0
    cleanedCSVFieldsIndex = 0
    Do While csvFieldIndex <= UBound(csvFields)
        cleanedField = csvFields(csvFieldIndex)
        If InStr(cleanedField, Chr(34)) And (csvFieldIndex < UBound(csvFields)) Then
            ' double quote detected in csvField and it's not the last field in the list
            concatFieldIndex = csvFieldIndex
            Do
                concatFieldIndex = concatFieldIndex + 1
                cleanedField = cleanedField & "," & csvFields(concatFieldIndex)
            Loop Until InStr(csvFields(concatFieldIndex), Chr(34)) Or (concatFieldIndex = UBound(csvFields))
            csvFieldIndex = concatFieldIndex
        End If
        csvFieldIndex = csvFieldIndex + 1
        
        If cleanedCSVFieldsIndex > 0 Then
            ReDim Preserve cleanedCSVFields(cleanedCSVFieldsIndex)
        End If
        cleanedCSVFields(cleanedCSVFieldsIndex) = cleanedField
        cleanedCSVFieldsIndex = cleanedCSVFieldsIndex + 1
    Loop
    
    CleanSplit = cleanedCSVFields
End Function
Public Function LINZCSVFilter(fieldName, rng, seedFilter) As String
    ' Creates a SQL "IN" filter for the given field name by comma seperating the values in cell rng
    ' fieldName is a field name belonging to a LINZ table
    ' rng is a selected range of cells
    ' seedFilter is a further restriction to apply to the generated filter
    
    Dim filter As String: filter = seedFilter
    If filter <> "" Then
        filter = filter + "%20AND%20"
    End If
    filter = filter + fieldName + "%20IN%20("
    
    Dim counter As Integer: counter = 0
    For Each cell In rng
        If IsEmpty(cell.Value) Then
            Exit For
        End If
        If counter > 0 Then
            filter = filter + ","
        End If
        filter = filter + "%27" + URLEncode(cell.Value) + "%27"
        counter = counter + 1
    Next cell
    LINZCSVFilter = filter + ")"
End Function
Public Function LINZCSV(typeName As String, filter As String) As String
    ' Return results as CSV string
    
    Dim query As String: query = "https://data.linz.govt.nz/services;key=" + key + "/wfs?service=WFS&version=2.0.0&request=GetFeature&typeNames=" + typeName + "&cql_filter=" + filter + "&outputFormat=CSV"
    
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", query, False
    http.send
    LINZCSV = http.responseText

End Function
Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String
  ' This function from http://stackoverflow.com/a/218199

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

