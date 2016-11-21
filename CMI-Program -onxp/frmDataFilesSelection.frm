VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataFilesSelection 
   Caption         =   "CMI-Analyser (Data Files Selection)"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   OleObjectBlob   =   "frmDataFilesSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataFilesSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************************************
'Program by: Somoud Saqfelhait - Touchstone
'07/06/2009
'***************************************************************************



Private Sub CommandButton4_Click()

End Sub

Private Sub cmdOK_Click()
    Unload Me
    frmMain.Show
End Sub



Private Sub UserForm_Activate()
    mpg.Value = Right(Year(Date) - 2005, 2)
    For i = Year(Date) + 1 To 2018
       mpg.Pages("pg" & i).Enabled = False
    Next i
        ThisWorkbook.Activate
        txtLocal05.Text = ThisWorkbook.Sheets(1).Range("B1").Value
        txtLocal06.Text = ThisWorkbook.Sheets(1).Range("B4").Value
        txtLocal07.Text = ThisWorkbook.Sheets(1).Range("B7").Value
        txtLocal08.Text = ThisWorkbook.Sheets(1).Range("B10").Value
        txtLocal09.Text = ThisWorkbook.Sheets(1).Range("B13").Value
        
        txtNational05.Text = ThisWorkbook.Sheets(1).Range("B2").Value
        txtNational06.Text = ThisWorkbook.Sheets(1).Range("B5").Value
        txtNational07.Text = ThisWorkbook.Sheets(1).Range("B8").Value
        txtNational08.Text = ThisWorkbook.Sheets(1).Range("B11").Value
        txtNational09.Text = ThisWorkbook.Sheets(1).Range("B14").Value
        
        txtRegional05.Text = ThisWorkbook.Sheets(1).Range("B3").Value
        txtRegional06.Text = ThisWorkbook.Sheets(1).Range("B6").Value
        txtRegional07.Text = ThisWorkbook.Sheets(1).Range("B9").Value
        txtRegional08.Text = ThisWorkbook.Sheets(1).Range("B12").Value
        txtRegional09.Text = ThisWorkbook.Sheets(1).Range("B15").Value
        
End Sub
'*************************2005*********************************************
Private Sub cmdLocal05_Click()
    txtLocal05.Text = getFileName("Select Local CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B1").Value = txtLocal05.Text
End Sub

Private Sub cmdNational05_Click()
    txtNational05.Text = getFileName("Select National Statistics Excel File:")
    ThisWorkbook.Sheets(1).Range("B2").Value = txtNational05.Text
End Sub

Private Sub cmdRegional05_Click()
    txtRegional05.Text = getFileName("Select Regional CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B3").Value = txtRegional05.Text
End Sub
'*************************2006*********************************************

Private Sub cmdLocal06_Click()
    txtLocal06.Text = getFileName("Select Local CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B4").Value = txtLocal06.Text
End Sub

Private Sub cmdNational06_Click()
    txtNational06.Text = getFileName("Select National Statistics Excel File:")
    ThisWorkbook.Sheets(1).Range("B5").Value = txtNational06.Text
End Sub

Private Sub cmdRegional06_Click()
    txtRegional06.Text = getFileName("Select Regional CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B6").Value = txtRegional06.Text
End Sub
'*************************2007*********************************************

Private Sub cmdLocal07_Click()
    txtLocal07.Text = getFileName("Select Local CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B7").Value = txtLocal07.Text
End Sub

Private Sub cmdNational07_Click()
    txtNational07.Text = getFileName("Select National Statistics Excel File:")
    ThisWorkbook.Sheets(1).Range("B8").Value = txtNational07.Text
End Sub

Private Sub cmdRegional07_Click()
    txtRegional07.Text = getFileName("Select Regional CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B9").Value = txtRegional07.Text
End Sub

'*************************2008*********************************************

Private Sub cmdLocal08_Click()
    txtLocal08.Text = getFileName("Select Local CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B10").Value = txtLocal08.Text
End Sub

Private Sub cmdNational08_Click()
    txtNational08.Text = getFileName("Select National Statistics Excel File:")
    ThisWorkbook.Sheets(1).Range("B11").Value = txtNational08.Text
End Sub

Private Sub cmdRegional08_Click()
    txtRegional08.Text = getFileName("Select Regional CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B12").Value = txtRegional08.Text
End Sub
'*************************2009*********************************************

Private Sub cmdLocal09_Click()
    txtLocal09.Text = getFileName("Select Local CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B13").Value = txtLocal09.Text
End Sub

Private Sub cmdNational09_Click()
    txtNational09.Text = getFileName("Select National Statistics Excel File:")
    ThisWorkbook.Sheets(1).Range("B14").Value = txtNational09.Text
End Sub

Private Sub cmdRegional09_Click()
    txtRegional09.Text = getFileName("Select Regional CMI Excel File:")
    ThisWorkbook.Sheets(1).Range("B15").Value = txtRegional09.Text
End Sub
