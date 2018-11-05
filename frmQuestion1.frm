VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuestion1 
   Caption         =   "Browser for input file..."
   ClientHeight    =   1701
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7252
   OleObjectBlob   =   "frmQuestion1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuestion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Question 1
'
' Revision History
' Date       Developer       Description
' 03-Nov-18  M Gore          Initial Version
'
' Create a form with three buttons, a text box and a label.
' a. One button should be labeled Browse.  When clicked, use the
'    Microsoft Common Dialog control to display a list of
'    folders/files.  After the user selects a file and clicks OK,
'    display the full path and filename in the text box.
' b. One button should be labeled OK.  When clicked, it should
'    display a message box with the filename.  If the user did
'    not choose a file, display a message box that indicates so.
' c. One button should be labeled Cancel.  When clicked, exit
'    the form.
'*************************************************************

Private Sub btnBrowse_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' btnBrowse_Click
' Uses the Excel GetOpenFileName method.  When the
' user clicks Open, display the full path
' and filename in tbxFilePath.  If no file is
' is chosen GetOpenFileName returns False, clear
' the contents of the file path text box, tbxFilePath.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sFile = Application.GetOpenFilename
    
    If sFile = "False" Then
        tbxFilePath.Text = ""
    Else
        tbxFilePath.Text = sFile
    End If
End Sub

Private Sub btnCancel_Click()
'''''''''''''''''''''''''''''
' btnCancel_Click
' Hide the form.
'''''''''''''''''''''''''''''
    frmQuestion1.Hide
End Sub

Private Sub btnOK_Click()
'''''''''''''''''''''''''''''''''''''''
' btnOK_Click
' Displays a message box with the
' filename, in other words the last
' segment of the full file path.
' Unless no file is chosen
' and the text property is null, then
' display a message box indicating so.
'''''''''''''''''''''''''''''''''''''''
    If tbxFilePath.Text = "" Then
        MsgBox "No file selected..."
    Else
        arrFPathSegments = Split(tbxFilePath.Text, "\")
        iFPSegSize = UBound(arrFPathSegments)
        
        MsgBox arrFPathSegments(iFPSegSize)
    End If
End Sub
