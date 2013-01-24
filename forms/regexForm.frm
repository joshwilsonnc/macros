VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} regexForm 
   Caption         =   "Regular Expression Find and Replace"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8730
   OleObjectBlob   =   "regexForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "regexForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub regexCancel_Click()

regexFind.Value = ""
regexForm.Hide

End Sub

Private Sub regexReplacement_Change()

If regexReplacement.Value = False Then
    regexReplace.Visible = False
    regexReplaceLabel.Visible = False
Else
    regexReplace.Visible = True
    regexReplaceLabel.Visible = True
End If

End Sub

Private Sub regexRun_Click()

regexForm.Hide

End Sub
