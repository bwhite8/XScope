Attribute VB_Name = "modXScope"
'
' Title:    modXScope
' Purpose:  Handles userform controls, builds and executes SQL queries, operates on ADODB recordsets, returns that data to the userform
' Related:  frmXScope
' Date:     1/5/2015
' Author:   Brian White
'




Sub runXScope(control As IRibbonControl)
'Sub runXScope()

    frmXScope.Show

End Sub

Function DetermineSQLX() As String

DetermineSQLX = _
"SELECT %REMOVED% " & _
"FROM %REMOVED% AS %REMOVED% " & _
"LEFT %REMOVED% AS %REMOVED% ON %REMOVED%=%REMOVED% AND %REMOVED%<>'%REMOVED%' " & _
"LEFT OUTER JOIN %REMOVED% AS %REMOVED% ON %REMOVED%=%REMOVED% AND %REMOVED%<>'%REMOVED%' " & _
"LEFT OUTER JOIN %REMOVED% AS %REMOVED% ON %REMOVED%=%REMOVED% AND %REMOVED%<>'%REMOVED%' " & _
"WHERE %REMOVED%=" & GetMA_ID & _
" AND" & _
btX & stX & a1X & a2X & a3X & cityX & stateX & zipX & " " & _
"WITH UR"

End Function

Function btX() As String

Dim billto As String

billto = frmXScope.txtBT.Value

If billto <> "" Then

    billto = UCase(billto)
    btX = " %REMOVED%='" & billto & "'"

Else

    btX = ""

End If

End Function
Function stX() As String

Dim shipto As String

shipto = frmXScope.txtST.Value

If shipto <> "" Then

    shipto = UCase(shipto)
    
    If btX <> "" Then
    
        stX = " AND %REMOVED% LIKE '" & shipto & "%'"
        
    Else
        
        stX = " %REMOVED% LIKE '" & shipto & "%'"
    
    End If
    
Else

    stX = ""

End If

End Function
Function a1X() As String

Dim address1 As String

address1 = frmXScope.txtA1.Value

If address1 <> "" Then

    address1 = UCase(address1)
    
    If btX <> "" Or stX <> "" Then
    
        a1X = " AND %REMOVED% LIKE '" & address1 & "%'"
    
    Else
    
        a1X = " %REMOVED% LIKE '" & address1 & "%'"
    
    End If
    
Else

    a1X = ""

End If

End Function
Function a2X() As String

Dim address2 As String

address2 = frmXScope.txtA2.Value

If address2 <> "" Then

    address2 = UCase(address2)
    
    If btX <> "" Or stX <> "" Or a1X <> "" Then
    
        a2X = " AND %REMOVED% LIKE '" & address2 & "%'"
        
    Else
    
        a2X = " %REMOVED% LIKE '" & address2 & "%'"
    
    End If
    
Else

    a2X = ""

End If

End Function
Function a3X() As String

Dim address3 As String

address3 = frmXScope.txtA3.Value

If address3 <> "" Then

    address3 = UCase(address3)
    
    If btX <> "" Or stX <> "" Or a1X <> "" Or a2X <> "" Then
        
        a3X = " AND %REMOVED% LIKE '" & address3 & "%'"
        
    Else
    
        a3X = " %REMOVED% LIKE '" & address3 & "%'"
    
    End If

Else

    a3X = ""

End If

End Function
Function cityX() As String

Dim city As String

city = frmXScope.txtCity.Value

If city <> "" Then

    city = UCase(city)
    
    If btX <> "" Or stX <> "" Or a1X <> "" Or a2X <> "" Or a3X <> "" Then
    
        cityX = " AND %REMOVED% LIKE '" & city & "%'"
    
    Else
    
        cityX = " %REMOVED% LIKE '" & city & "%'"
    
    End If
    
Else

    cityX = ""

End If

End Function
Function stateX() As String

Dim sstate As String

sstate = frmXScope.txtState.Value

If sstate <> "" Then

    sstate = UCase(sstate)
    
    If btX <> "" Or stX <> "" Or a1X <> "" Or a2X <> "" Or a3X <> "" Or cityX <> "" Then
    
        stateX = " AND %REMOVED%='" & sstate & "'"
        
    Else
    
        stateX = " %REMOVED%='" & sstate & "'"
        
    End If
    
Else

    stateX = ""

End If

End Function

Function zipX() As String

Dim zip As String

zip = frmXScope.txtZip.Value

If zip <> "" Then

    If btX <> "" Or stX <> "" Or a1X <> "" Or a2X <> "" Or a3X <> "" Or cityX <> "" Or stX <> "" Then
    
        zipX = " AND %REMOVED%='" & zip & "'"
    
    Else
    
        zipX = " %REMOVED%='" & zip & "'"
    
    End If
    
Else

    zipX = ""

End If

End Function


Function GetMA_ID() As String

    Dim strSQLM, temp As String
    
    Dim maList As New ADODB.Recordset
    Dim cnn1 As New ADODB.Connection
    
    strSQLM = "SELECT %REMOVED% FROM %REMOVED% WHERE %REMOVED%='" & getFullMasterX & "' AND %REMOVED%='%REMOVED%' AND %REMOVED%<>'%REMOVED%' WITH UR"
    cnn1.Provider = "%REMOVED%"
    
    cs = "Server=srv:%REMOVED%;data source=%REMOVED%;UID=%REMOVED%;PWD=%REMOVED%;Connect Timeout=10"
    
    On Error Resume Next
    cnn1.Open cs
    
    If cnn1.state <> adStateOpen Then
    
        MsgBox ("Couldn't connect to %REMOVED%.  Check your connection settings.")
        Exit Function
        
    End If
    
    Set maList = New ADODB.Recordset
    
    On Error GoTo ErrorHandler3:
    maList.Open Source:=strSQLM, ActiveConnection:=cnn1, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    
    GetMA_ID = maList.GetString()
    GetMA_ID = Replace(GetMA_ID, Chr(13), "")
    
    maList.Close
    cnn1.Close
    
    Set cnn1 = Nothing
    Set maList = Nothing
    
    Exit Function
ErrorHandler3:
    
        cnn.Close
        
        Set cnn1 = Nothing
        
        frmXScope.lblStatus.Caption = "Error.  Ensure Master provided is correct."
        

End Function

Function getFullMasterX() As String
Dim sMaster As String
Dim sDiv As String

sMaster = frmXScope.txtMaster.Value
sDiv = frmXScope.cbDiv.Value

If Len(sMaster) < 5 Then
    
    getFullMasterX = "000000" & sMaster & "001" & sDiv
    
ElseIf Len(sMaster) < 6 Then
    
    getFullMasterX = "00000" & sMaster & "001" & sDiv
    
ElseIf Len(sMaster) < 7 Then
    
    getFullMasterX = "0000" & sMaster & "001" & sDiv
    
ElseIf Len(sMaster) < 8 Then
        
    getFullMasterX = "000" & sMaster & "001" & sDiv
    
ElseIf Len(sMaster) < 9 Then
    
    getFullMasterX = "00" & sMaster & "001" & sDiv
    
ElseIf Len(sMaster) < 10 Then
        
    getFullMasterX = "0" & sMaster & "001" & sDiv
     
ElseIf Len(sMaster) < 11 Then
    
    getFullMasterX = sMaster & "001" & sDiv
        
Else
        
    getFullMasterX = sMaster & sDiv
    
End If



End Function
Sub GetResults()

Dim cs, strSQL As String
Dim resultCount, n As Integer

strSQL = DetermineSQLX

Dim recResults As New ADODB.Recordset
Dim cnn As New ADODB.Connection

cnn.Provider = "%REMOVED%"

cs = "Server=srv:%REMOVED%;data source=%REMOVED%;UID=%REMOVED%;PWD=%REMOVED%;Connect Timeout=10"

On Error Resume Next
cnn.Open cs

If cnn.state <> adStateOpen Then

    MsgBox ("Couldn't connect to %REMOVED%.  Check your connection settings.")
    Exit Sub
    
End If

Set recResults = New ADODB.Recordset

On Error GoTo ErrorHandler3:
recResults.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic

resultCount = recResults.RecordCount

If Not resultCount > 0 Then
    GoTo ErrorHandler2
End If

If resultCount > 20000 Then
    
    MsgBox ("The number of records exceeds 20K.  Please pull a %REMOVED% using the %REMOVED% Query Tool or narrow down your search results by incorporating additional filters.")
    Exit Sub

End If

n = 0
recResults.MoveFirst
Do

    With frmXScope.listResults
        .AddItem
        .List(n, 0) = recResults.Fields(0).Value
        .List(n, 1) = recResults.Fields(1).Value
        .List(n, 2) = recResults.Fields(2).Value
        .List(n, 3) = recResults.Fields(3).Value
        .List(n, 4) = recResults.Fields(4).Value
        
        If recResults.Fields(5).Value <> "" Then
            .List(n, 5) = recResults.Fields(5).Value
        Else
            .List(n, 5) = vbTab
        End If
        
        If recResults.Fields(6).Value <> "" Then
            .List(n, 6) = recResults.Fields(6).Value
        Else
            .List(n, 6) = vbTab
        End If
        
        .List(n, 7) = recResults.Fields(7).Value
        .List(n, 8) = recResults.Fields(8).Value
        .List(n, 9) = recResults.Fields(9).Value
    End With
    n = n + 1
    recResults.MoveNext

Loop Until recResults.EOF

recResults.Close
cnn.Close

Set recResults = Nothing
Set cnn = Nothing

frmXScope.lblStatus.Caption = resultCount & " Ship-To Locations Found"

MsgBox ("Successful Execution.  Click OK to continue.")

Exit Sub

ErrorHandler2:
    
    recResults.Close
    cnn.Close
    
    Set recResults = Nothing
    Set cnn = Nothing
    
    frmXScope.lblStatus.Caption = "No Results"

    Exit Sub
    
ErrorHandler3:
    
    cnn.Close
    
    Set cnn = Nothing
    
    frmXScope.lblStatus.Caption = "Error! Ensure the Master and Div provided are correct."

End Sub

' Sub SortListBox written primarily by: postman2000
' Retrieved From: http://www.ozgrid.com/forum/showthread.php?t=71509
' Edited by Brian White 1/5/2015

Sub SortListBox(oLb As MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
    Dim vaItems As Variant
    Dim i As Long, j As Long
    Dim c As Integer
    Dim vTemp As Variant
     
     
    If oLb.ListCount = 0 Then
        Exit Sub
    End If
    
    'Put the items in a variant array
    vaItems = oLb.List
     
    
     'Sort the Array Alphabetically(1)
    If sType = 1 Then
    
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                 'Sort Ascending (1)
                If sDir = 1 Then
                    If vaItems(i, sCol) > vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                     
                     'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If vaItems(i, sCol) < vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
                 
            Next j
        Next i
         'Sort the Array Numerically(2)
         '(Substitute CInt with another conversion type (CLng, CDec, etc.) depending on type of numbers in the column)
    ElseIf sType = 2 Then
    
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                 'Sort Ascending (1)
                If sDir = 1 Then
                    If CInt(vaItems(i, sCol)) > CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                     
                     'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
                 
            Next j
        Next i
    End If
     
     'Set the list to the array
    oLb.List = vaItems

    Set vaItems = Nothing
    
End Sub

Sub ExpandST()
Dim cs, strSQL, strST As String
Dim resultCount, n, selRow As Integer

On Error GoTo ErrorHandler:
selRow = frmXScope.listResults.ListIndex
strST = frmXScope.listResults.List(selRow, 2)
strSTAT = frmXScope.listResults.List(selRow, 0)

strSQL = "SELECT %REMOVED% " & _
"FROM %REMOVED% AS %REMOVED% " & _
"LEFT OUTER JOIN %REMOVED% AS %REMOVED% ON %REMOVED%=%REMOVED% AND %REMOVED%<>'%REMOVED%' " & _
"LEFT OUTER JOIN %REMOVED% AS %REMOVED% ON %REMOVED%=%REMOVED% AND %REMOVED%<>'%REMOVED%' " & _
"WHERE %REMOVED%=" & GetMA_ID & _
" AND" & _
" %REMOVED%='" & strST & "' " & _
"WITH UR"

Dim recResults As New ADODB.Recordset
Dim cnn As New ADODB.Connection

cnn.Provider = "%REMOVED%"

cs = "Server=srv:%REMOVED%;data source=%REMOVED%;UID=%REMOVED%;PWD=%REMOVED%;Connect Timeout=10"

On Error Resume Next
cnn.Open cs

If cnn.state <> adStateOpen Then

    MsgBox ("Couldn't connect to Datamart.  Check your connection settings.")
    Exit Sub
    
End If

Set recResults = New ADODB.Recordset

On Error GoTo ErrorHandler3:
recResults.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic

resultCount = recResults.RecordCount

If Not resultCount > 0 Then
    GoTo ErrorHandler2
End If

    bcType = recResults.Fields(0).Value
    bcLevel = recResults.Fields(1).Value

recResults.Close
cnn.Close

Set recResults = Nothing
Set cnn = Nothing

If strSTAT = "%REMOVED%" Then

    strSTAT = "Active"

ElseIf strSTAT = "%REMOVED%" Then

    strSTAT = "Deleted"

End If

If bcType = "%REMOVED%" Then

    bcType = "Free Format"

ElseIf bcType = "%REMOVED%" Or bcType = "%REMOVED%" Then

    bcType = "Required"

End If

If bcLevel = "%REMOVED%" Then
    
    bcLevel = "Master"

ElseIf bcLevel = "%REMOVED%" Then
    
    bcLevel = "Bill-To"

ElseIf bcLevel = "%REMOVED%" Then

    bcLevel = "Ship-To"

End If

frmXScope.lblStatus.Caption = "Ship-To ID " & Trim(strST) & " is " & strSTAT & ". " & "Budget Center Type: " & bcType & " at the " & bcLevel & " Level."

Exit Sub
ErrorHandler:
    Exit Sub
    
ErrorHandler2:
    
    recResults.Close
    cnn.Close
    
    Set recResults = Nothing
    Set cnn = Nothing

    Exit Sub
    
ErrorHandler3:
    
    cnn.Close
    Set cnn = Nothing
    
End Sub
