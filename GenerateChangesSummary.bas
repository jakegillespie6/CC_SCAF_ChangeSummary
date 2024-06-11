Attribute VB_Name = "Module2"
Sub Create_Email_Comments()

    Dim rng As Range
    
    Dim oppName As String
    Dim oppID As String
    oppName = ActiveWorkbook.ActiveSheet.Range("B4")
    oppID = ActiveWorkbook.ActiveSheet.Range("B3")

    
    Dim tFiberMileage1 As String
    tFiberMileage1 = ActiveWorkbook.ActiveSheet.Range("B7")
    Dim tFiberMileage2 As String
    tFiberMileage2 = ActiveWorkbook.ActiveSheet.Range("C7")
    Dim tFiberMileage As String
    
    Dim avgRANMRC1 As String
    avgRANMRC1 = ActiveWorkbook.ActiveSheet.Range("B8")
    Dim avgRANMRC2 As String
    avgRANMRC2 = ActiveWorkbook.ActiveSheet.Range("C8")
    Dim avgRANMRC As String
    
    Dim avgFiberMRC1 As String
    avgFiberMRC1 = ActiveWorkbook.ActiveSheet.Range("B9")
    Dim avgFiberMRC2 As String
    avgFiberMRC2 = ActiveWorkbook.ActiveSheet.Range("C9")
    Dim avgFiberMRC As String
    
    Dim RANDensity1 As String
    RANDensity1 = ActiveWorkbook.ActiveSheet.Range("B10")
    Dim RANDensity2 As String
    RANDensity2 = ActiveWorkbook.ActiveSheet.Range("C10")
    Dim RANDensity As String
    
    Dim numZLOC1 As String
    numZLOC1 = ActiveWorkbook.ActiveSheet.Range("B6")
    Dim numZLOC2 As String
    numZLOC2 = ActiveWorkbook.ActiveSheet.Range("C6")
    Dim numZLOC As String
    
    Dim wsSCAF1 As Worksheet
    Dim wsSCAF2 As Worksheet

    Dim tblSCAF1 As Excel.ListObject
    Dim tblSCAF2 As Excel.ListObject
    
    Set wsSCAF1 = ThisWorkbook.Worksheets("First SCAF Site Config App")
    Set wsSCAF2 = ThisWorkbook.Worksheets("Second SCAF Site Config App")
    Set tblSCAF1 = ThisWorkbook.Worksheets("First SCAF Site Config App").ListObjects("First_SCAF_Site_Config_App")
    Set tblSCAF2 = ThisWorkbook.Worksheets("Second SCAF Site Config App").ListObjects("Second_SCAF_Site_Config_App")
    
    Dim SCAF1HubDict As Scripting.Dictionary
    Dim SCAF2HubDict As Scripting.Dictionary
    Set SCAF1HubDict = New Scripting.Dictionary
    Set SCAF1HubDict = New Scripting.Dictionary
    Set SCAF1HubDict = Get_Nodes_Per_Hub(wsSCAF1, tblSCAF1)
    Set SCAF2HubDict = Get_Nodes_Per_Hub(wsSCAF2, tblSCAF2)

    Dim SCAF1PoleDict As Scripting.Dictionary
    Dim SCAF2PoleDict As Scripting.Dictionary
    Set SCAF1PoleDict = New Scripting.Dictionary
    Set SCAF2PoleDict = New Scripting.Dictionary
    Set SCAF1PoleDict = Get_Nodes_Per_Pole(wsSCAF1, tblSCAF1)
    Set SCAF2PoleDict = Get_Nodes_Per_Pole(wsSCAF2, tblSCAF2)

    Dim olApp As Outlook.Application
    Dim olEmail As Outlook.MailItem
    Set olApp = New Outlook.Application
    Set olEmail = olApp.CreateItem(olMailItem)




'Change in parent values
    tFiberMileage = "Total Fiber Mileage: " & Val_Change(tFiberMileage1, tFiberMileage2)
    avgRANMRC = "Average RAN MRC: " & Val_Change(avgRANMRC1, avgRANMRC2)
    avgFiberMRC = "Average Fiber MRC: " & Val_Change(avgFiberMRC1, avgFiberMRC2)
    RANDensity = "RAN Density*: " & Val_Change(RANDensity1, RANDensity2)
    numZLOC = "Node Count: " & Val_Change(numZLOC1, numZLOC2)


Dim dictChange As String
dictChange = Hub_Dict_Change(SCAF1HubDict, SCAF2HubDict)
Dim poleChange As String
poleChange = Pole_Dict_Change(SCAF1PoleDict, SCAF2PoleDict)

    


With olEmail
.BodyFormat = olFormatHTML
.Display
.HTMLBody = "Hello Team," & "<br>" & "Attach are the following opportunities ready to go to the customer for review and approval, this is .  Please see notes per attach Opportunity SCAF, let us if you have any questions or concerns." _
            & "<br><br>" & "<b>" & "Opportunity Name:  " & oppName _
            & "<br>" & "Opportunity ID:  " & oppID & "</b>" _
            & "<br><br>" & tFiberMileage _
            & "<br>" & avgRANMRC _
            & "<br>" & avgFiberMRC _
            & "<br>" & RANDensity _
            & "<br>" & numZLOC _
            & "<br>" & dictChange _
            & "<br>" & poleChange
            
            
                        
.To = ""
.CC = ""
.Subject = "SCAF Changes for " & oppName
    
End With
End Sub
Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in

    Set TempWB = Workbooks.Add(1)

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Public Function Get_Nodes_Per_Hub(wsSCAF As Worksheet, tblSCAF As Excel.ListObject) As Scripting.Dictionary
    Dim Val As String
    Dim hubDict As Scripting.Dictionary
    Set hubDict = New Scripting.Dictionary

For Each lr In tblSCAF.ListRows
    Val = wsSCAF.Cells(lr.Index + 1, 2)
    If hubDict.Exists(Val) Then
        hubDict(Val) = hubDict(Val) + 1
    Else
        hubDict.Add Val, 1
    End If
Next lr

Set Get_Nodes_Per_Hub = hubDict
End Function

Function Val_Change(val1 As String, val2 As String) As String
    If val1 <> val2 Then
        Val_Change = val1 & " -> " & val2
    Else
        Val_Change = val1
    End If
End Function

Function Hub_Dict_Change(dict1 As Scripting.Dictionary, dict2 As Scripting.Dictionary) As String
    Dim Str As String
    Str = "<br><br><b> HUBS (" & dict2.Count() & ")</b>"
    For Each Key In dict2
        If dict1.Exists(Key) Then
            Str = Str & "<br>" & vbTab & "    - " & Key & ": " & dict2(Key)
        Else
            Str = Str & "<br>" & vbTab & "    - " & Key & " (NEW): " & dict2(Key)
        End If
    Next Key
    Hub_Dict_Change = Str
End Function

Function Get_Nodes_Per_Pole(wsSCAF As Worksheet, tblSCAF As Excel.ListObject) As Scripting.Dictionary
    Dim Val As String
    Dim poleDict As Scripting.Dictionary
    Set poleDict = New Scripting.Dictionary
    
    For Each lr In tblSCAF.ListRows
        Val = wsSCAF.Cells(lr.Index + 1, 12)
        If poleDict.Exists(Val) Then
            poleDict(Val) = poleDict(Val) + 1
        Else
            poleDict.Add Val, 1
        End If
    Next lr
    Set Get_Nodes_Per_Pole = poleDict
End Function

Function Pole_Dict_Change(dict1 As Scripting.Dictionary, dict2 As Scripting.Dictionary) As String
    Dim Str As String
    Str = "<br><br><b> Nodes By Pole Type</b>"
    
    For Each Key In dict1
        If dict2.Exists(Key) = False Then
            Str = Str & "<br>" & vbTab & "    - " & Key & ": " & dict1(Key) & " -> 0"
        End If
    Next Key
    
    For Each Key In dict2
        If dict1.Exists(Key) = False Then
            Str = Str & "<br>" & vbTab & "    - " & Key & " (NEW): " & dict2(Key)
        ElseIf dict1(Key) <> dict2(Key) Then
            Str = Str & "<br>" & vbTab & "    - " & Key & ": " & dict1(Key) & " -> " & dict2(Key)
        Else
            Str = Str & "<br>" & vbTab & "    - " & Key & ": " & dict2(Key)
        End If
    Next Key
    Pole_Dict_Change = Str
End Function
