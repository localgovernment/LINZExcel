' ************************************************************************************************
' * LINZExcel : A collection of Excel VBA functions that access the LINZ data.linz.govt.nz WFS API
' *
' * The MIT License (MIT)
' *
' * Copyright (c) 2014 Taupo District Council, Dion Liddell.
' ************************************************************************************************

' ******* Replace the following with your LINZ API Key *********
Public Const key As String = "my LINZ API key"
' **************************************************************

Public Function GetMainParcelID(valuation As String) As String
    ' Taupo District Council Mapi API - returns the main parcel ID for the given valuation number
    Dim query As String: query = "http://gis.taupodc.govt.nz/arcgis/rest/services/Mapi/TaupoProperty/MapServer/0/query?where=valuation_id+%3D+%27" + valuation + "%27&outFields=valuation_id%2C+m_parcel_id&returnGeometry=false&f=json"
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim xmlDoc As Object: Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    http.Open "GET", query, False
    http.send
    GetMainParcelID = ExtractJSONValue(http.responseText, "m_parcel_id")
End Function
Public Function ExtractJSONValue(json As String, key As String) As String
    Dim start As Integer: start = InStr(json, key)
    start = InStr(json, "features")
    start = InStr(start, json, key)
    start = start + Len(key) + 3
    Dim finish As Integer: finish = InStr(start, json, Chr(34))
    Dim length As Integer: length = finish - start
    ExtractJSONValue = Mid(json, start, length)
End Function
Public Function GetTitles(parcelID As String) As String
    ' Returns a list of (current) LINZ titles for the given ParcelID
    GetTitles = LINZ("layer-772", "id=" + parcelID, "titles")
End Function
Public Function GetEncumbrancees(title As String) As String
    ' Returns a list of (current) LINZ encumbrancees for the given title
    GetEncumbrancees = LINZ("table-1695", "title_no=%27" + FirstInList(title) + "%27%20AND%20current=%27true%27", "encumbrancees")
End Function
Public Function GetInstrumentNumbers(title As String) As String
    ' Returns a list of (current) LINZ Instrument numbers for the given title
    GetInstrumentNumbers = LINZ("table-1695", "title_no=%27" + FirstInList(title) + "%27%20AND%20current=%27true%27", "instrument_number")
End Function
Public Function GetInstrumentTypes(title As String) As String
    ' Returns a list of (current) LINZ instrument types for the given title – should be same order as instrument numbers
    GetInstrumentTypes = LINZ("table-1695", "title_no=%27" + FirstInList(title) + "%27%20AND%20current=%27true%27", "instrument_type")
End Function
Public Function GetSurnames(title As String) As String
    ' Returns a list of LINZ surnames for the given title
    GetSurnames = LINZ("table-1564", "title_no=%27" + FirstInList(title) + "%27", "prime_surname")
End Function
Public Function FirstInList(aList As String) As String
    Dim comma As Integer: comma = InStr(aList, ",")
    If comma = 0 Then
        FirstInList = aList
    Else
        FirstInList = Trim(Left(aList, comma - 1))
    End If
End Function
Public Sub ProcessValuations()
    ' Process all selected valuations - TDC specific but can be modified
    ' TODO: requires serious refactoring
    
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim rng As Range
    Dim parcelID As String
    Dim titles() As String
    Dim ttle As String
    Dim currow As Integer
    
    Set rng = Selection
    Set sourceSheet = ActiveSheet
    Set destSheet = Sheets.Add
    
    ' format columns as text and populate headings
    destSheet.Columns("A:G").NumberFormat = "@"
    
    destSheet.Range("A1").Value = "Valuation No."
    destSheet.Range("B1").Value = "Main Parcel ID"
    destSheet.Range("C1").Value = "LINZ Title"
    destSheet.Range("D1").Value = "LINZ Surnames"
    destSheet.Range("E1").Value = "Encumbrancee"
    destSheet.Range("F1").Value = "Instrument Numbers"
    destSheet.Range("G1").Value = "Instrument Types"
    
    ' get data from Mapi and LINZ
    For Each cell In rng
        If IsEmpty(cell.Value) Then
            Exit For
        End If
        parcelID = GetMainParcelID(cell.Value)
        titles = Split(GetTitles(parcelID), ",")
        For Each title In titles
            ttle = Trim(CStr(title))
            currow = currow + 1
            destSheet.Range("A1").Offset(currow, 0).Value = cell.Value
            destSheet.Range("B1").Offset(currow, 0).Value = parcelID
            destSheet.Range("C1").Offset(currow, 0).Value = ttle
            destSheet.Range("D1").Offset(currow, 0).Value = GetSurnames(ttle)
            destSheet.Range("E1").Offset(currow, 0).Value = GetEncumbrancees(ttle)
            destSheet.Range("F1").Offset(currow, 0).Value = GetInstrumentNumbers(ttle)
            destSheet.Range("G1").Offset(currow, 0).Value = GetInstrumentTypes(ttle)
        Next title
    Next cell

End Sub
Public Function LINZ(typeName As String, filter As String, element As String) As String
    ' using the following as reference
    ' http://libkod.info/officexml-CHP-9-SECT-5.shtml#officexml-CHP-9-EX-4
    ' http://www.wikihow.com/Create-a-User-Defined-Function-in-Microsoft-Excel
    ' http://stackoverflow.com/questions/11245733/declaring-early-bound-msxml-object-throws-an-error-in-vba
    ' http://stackoverflow.com/questions/19117667/how-to-read-xml-attributes-using-vba-to-excel
    ' http://stackoverflow.com/questions/5297068/read-xml-attribute-vba
    
    Dim query As String: query = "https://data.linz.govt.nz/services;key=" + key + "/wfs?service=WFS&version=2.0.0&request=GetFeature&typeNames=" + typeName + "&cql_filter=" + filter
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim xmlDoc As Object: Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Dim elements As Object
    Dim el As Variant
    Dim results As String: results = ""
    
    'create HTTP request to query URL - make sure to have
    'that last "False" there for synchronous operation
    http.Open "GET", query, False
    
    'send HTTP request
    http.send

    'parse result
    xmlDoc.LoadXML http.responseText
    
    'gather and return element text(s)
    Set elements = xmlDoc.getElementsByTagName("data.linz.govt.nz:" + element)
    For Each el In elements
        results = results + IIf(results = "", "", ", ")
        results = results + el.Text
    Next
    
    LINZ = results

End Function
