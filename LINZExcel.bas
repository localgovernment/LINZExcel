Attribute VB_Name = "LINZExcel"

' ************************************************************************************************
' * LINZExcel : A collection of Excel VBA functions that access the LINZ data.linz.govt.nz WFS API
' *
' * The MIT License (MIT)
' *
' * Copyright (c) 2014 Taupo District Council, Dion Liddell.
' ************************************************************************************************

' ******* Replace the following with your LINZ API Key *********
Public Const key as String = "my LINZ API key"
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
    Debug.Print json, key
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
Public Function FirstInList(aList As String) As String
    Dim comma As Integer: comma = InStr(aList, ",")
    If comma = 0 Then
        FirstInList = aList
    Else
        FirstInList = Trim(Left(aList, comma - 1))
    End If
End Function
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
