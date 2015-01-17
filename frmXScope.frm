VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXScope 
   Caption         =   "XScope Tool"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   23460
   OleObjectBlob   =   "frmXScope.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmXScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Title:    frmXScope
' Purpose:  Handles userform controls
' Related:  modXScope
' Date:     01-05-2015
' Author:   Brian White
'


Private Sub imgCopyST_Click()

Dim objCopy As MSForms.DataObject

Set objCopy = New MSForms.DataObject

Dim selRow As Integer
Dim strST As String

strST = ""

    On Error GoTo ErrorHandler2:
    selRow = frmXScope.listResults.ListIndex
    
    On Error GoTo ErrorHandler2:
    strST = frmXScope.listResults.List(selRow, 2)
    
    On Error GoTo ErrorHandler2:
    strST = Trim(Replace(strST, Chr(13), ""))

On Error GoTo ErrorHandler:
objCopy.SetText strST
objCopy.PutInClipboard

Exit Sub
ErrorHandler:

MsgBox ("Function Not Supported")

ErrorHandler2:
    
End Sub

Private Sub lblExpandST_Click()

    modXScope.ExpandST
    
End Sub

Private Sub lblSortA1_Click()
 
    On Error GoTo ErrorNum:
    modXScope.SortListBox listResults, 4, 1, 1
    Exit Sub

ErrorNum:
    On Error Resume Next
    modXScope.SortListBox listResults, 4, 2, 1
    
End Sub

Private Sub lblSortA2_Click()

    On Error GoTo ErrorNum:
    modXScope.SortListBox listResults, 5, 1, 1
    Exit Sub
    
ErrorNum:
    On Error Resume Next
    modXScope.SortListBox listResults, 5, 2, 1
    
End Sub

Private Sub lblSortA3_Click()

    On Error GoTo ErrorNum:
    modXScope.SortListBox listResults, 6, 1, 1
    Exit Sub

ErrorNum:
    On Error Resume Next
    modXScope.SortListBox listResults, 6, 2, 1
    Exit Sub
    
End Sub

Private Sub lblSortBT_Click()
    
    On Error GoTo ErrorNum:
    modXScope.SortListBox listResults, 1, 1, 1
    Exit Sub
    
ErrorNum:
    On Error Resume Next
    modXScope.SortListBox listResults, 1, 2, 1
    
End Sub

Private Sub lblSortCity_Click()

    On Error Resume Next
    modXScope.SortListBox listResults, 7, 1, 1

End Sub

Private Sub lblSortCo_Click()
    
    On Error GoTo ErrorNum:
    modXScope.SortListBox listResults, 3, 1, 1
    Exit Sub

ErrorNum:
    On Error Resume Next
    modXScope.SortListBox listResults, 3, 2, 1

    
End Sub

Private Sub lblSortST_Click()
    
    On Error GoTo ErrorNum:
    modXScope.SortListBox listResults, 2, 1, 1
    Exit Sub
    
ErrorNum:
    On Error Resume Next
    modXScope.SortListBox listResults, 2, 2, 1
    
End Sub

Private Sub lblSortState_Click()
    
    On Error Resume Next
    modXScope.SortListBox listResults, 8, 1, 1

End Sub

Private Sub lblSortStatus_Click()
    
    On Error Resume Next
    modXScope.SortListBox listResults, 0, 1, 1
    
End Sub

Private Sub lblSortZip_Click()
    
    On Error GoTo ErrorAlpha:
    modXScope.SortListBox listResults, 9, 1, 1
    Exit Sub

ErrorAlpha:
    On Error Resume Next
    modXScope.SortListBox listResults, 9, 2, 1
    
End Sub

Private Sub ToggleButton1_Click()

    If ToggleButton1.Value = True Then
        
        Me.Height = 75
        Me.Width = 50
    
    ElseIf ToggleButton1.Value = False Then
        
        Me.Height = 344
        Me.Width = 1177
        
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    txtMaster.SetFocus
    cbDiv.AddItem "ATL"
    cbDiv.AddItem "BOS"
    cbDiv.AddItem "CHI"
    cbDiv.AddItem "DAL"
    cbDiv.AddItem "DC"
    cbDiv.AddItem "DET"
    cbDiv.AddItem "HOU"
    cbDiv.AddItem "LA"
    cbDiv.AddItem "NAT"
    cbDiv.AddItem "NYC"
    cbDiv.AddItem "PHL"
    cbDiv.AddItem "RCH"

End Sub

Private Sub cmdSearchX_Click()

listResults.Clear

Application.ScreenUpdating = False

If txtMaster.Value <> "" And cbDiv.Value <> "" Then
    
    lblStatus.Visible = True
    
    lblStatus.Caption = "Please Wait. Your Request is Being Processed..."

    Application.Wait (Now + TimeValue("0:00:2"))
    
    'listResults.Text = modXScope.GetResults
    
    modXScope.GetResults
    
Else
    
    lblStatus.Caption = "Please enter a Master and Div."

End If

Application.ScreenUpdating = True

End Sub
