VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLocal 
   Caption         =   "Local Data"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill Range cells with 0 if  they are empty"
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Frame fraRefrance 
      Caption         =   "DataSet Coding"
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   10575
      Begin VB.CommandButton cmdSelectD 
         Caption         =   "&Select"
         Height          =   375
         Left            =   7320
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtDataD 
         Height          =   375
         Left            =   5760
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtSheetD 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowseD 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   9240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Range:"
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Worksheet Name:"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label t 
         Caption         =   "Data Set Codes File name and Path"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Frame fraRaw 
      Caption         =   "Local Data File"
      Height          =   2415
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   10575
      Begin VB.CommandButton cmdSelectE 
         Caption         =   "Sele&ct"
         Height          =   375
         Left            =   7920
         TabIndex        =   22
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtRawR 
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtSheetR 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtSheetR 
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtRawR 
         Height          =   375
         Index           =   0
         Left            =   6240
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowseL 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   9240
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtRaw 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Local Data Range:"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   21
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Worksheet Name:"
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Worksheet Name:"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Ethnicity Range:"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblRaw 
         Caption         =   "Local Raw Data File name and Path"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oExcel As Excel.Application
Dim oWBL, oWBD, oWBR As Excel.Workbook
Dim oWSL, oWSD As Excel.Worksheet
Dim rRangeL, rRangeD, rRangeDE As Excel.Range
 
Private Sub Common_Error()
'---------------------------------
'This is used to catch the error
'that occurs when Cancel is clicked,
'otherwise the problem will crash.
'---------------------------------

'If error is 32755("Cancel was clicked") then
    If Err.Number = 32755 Then
        'Do nothing because
        'nothing is required...
         Set oWSL = Nothing
    
        If Not oWBL Is Nothing Then oWBL.Close False
     
        Set oWBL = Nothing
      
        oExcel.Quit
        Set oExcel = Nothing
    End If

End Sub

Private Sub cmdBrowseD_Click()
        '-------------------------------------
    'The Open command button is clicked.
    '-------------------------------------
    
    'If an error occurs goto ErrorH(below)...
    On Error GoTo ErrorH
    
    'This setups the dialog box to only
    'allow Icons to be loaded...
    CommonDialog.Filter = "Excel File|*.xls|*.xlsx"
    
    'The Open dialog box is shown...
    CommonDialog.ShowOpen
    txtData.Text = CommonDialog.FileName
    'If an error occurs(usually when Cancel is clicked)
   


 
    'oExcel.Visible = False
    On Error GoTo 0

        Application.DisplayAlerts = True

    
ErrorH:
       
    Common_Error 'Call the Common_Error sub-routine...
    Exit Sub         'Exit the sub(if an error)...

    
End Sub

Private Sub cmdBrowseL_Click()
        '-------------------------------------
    'The Open command button is clicked.
    '-------------------------------------
    
    'If an error occurs goto ErrorH(below)...
    On Error GoTo ErrorH
    
    'This setups the dialog box to only
    'allow Icons to be loaded...
    CommonDialog.Filter = "Excel File|*.xls"
    
    'The Open dialog box is shown...
    CommonDialog.ShowOpen
    txtRaw.Text = CommonDialog.FileName
    'If an error occurs(usually when Cancel is clicked)
    On Error GoTo 0
    Application.DisplayAlerts = True
ErrorH:
    Common_Error 'Call the Common_Error sub-routine...
    Exit Sub         'Exit the sub(if an error)...

    
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFill_Click()
Dim i, j As Integer
    Set oWBL = oExcel.Workbooks.Open(txtRaw.Text, False)
bb:     Set rRangeL = oWBL.Application.InputBox(Prompt:="Please select range with your Mouse.", Title:="SPECIFY RANGE", Type:=8)
     
     For i = 1 To rRangeL.Rows.Count 'selected codes e.g language
        For j = 1 To rRangeL.Columns.Count
                If rRangeL.Cells(i, j).Value = 0 Then
                    rRangeL.Cells(i, j).Value = 0
                End If
            Next j
        Next i
      GoTo bb
    On Error GoTo ss

        'If Not oWBL Is Nothing Then oWBL.Close True
  
        'Set oWBL = Nothing

ss:         Exit Sub
End Sub

Private Sub cmdGenerate_Click()
Dim k, i, j As Integer
Set oWBR = oExcel.Workbooks.Add
Dim sum As Integer

    For k = 1 To rRangeD.Rows.Count 'selected codes e.g language
    oWBR.Sheets("Sheet1").Cells(1, k).Value = rRangeD(k, 1).Value
        For i = 1 To 17  'ethnicity
            sum = 0
            For j = 1 To rRangeDE.Rows.Count
                If rRangeDE(j, 1).Value = oWBD.Sheets(1).Cells(i, 3).Value _
                And rRangeL(j, 1).Value = rRangeD(k, 1).Value Then
                    sum = sum + 1
                End If
            Next j
            oWBR.Sheets("Sheet1").Cells(i + 1, k).Value = sum
        Next i
    Next k
End Sub

Private Sub cmdSelectE_Click()
    
    Set oWBL = oExcel.Workbooks.Open(txtRaw.Text, False)
     Set rRangeL = oWBL.Application.InputBox(Prompt:="Please select ethnicity range with your Mouse.", Title:="SPECIFY RANGE", Type:=8)
    
    'Set oWSL = oWBD.ActiveSheet
    txtSheetR(0).Text = oWBL.ActiveSheet.Name
    txtRawR(0).Text = rRangeL.Address
    'oExcel.Visible = False
    Set rRangeDE = oWBL.Application.InputBox(Prompt:="Please select a range with your Mouse.", Title:="SPECIFY RANGE", Type:=8)
    txtSheetR(1).Text = oWBL.ActiveSheet.Name
    txtRawR(1).Text = rRangeDE.Address
End Sub

Private Sub cmdSelectD_Click()
    Set oWBD = oExcel.Workbooks.Open(txtData.Text, False)
    Set rRangeD = oWBD.Application.InputBox(Prompt:="Please select a range with your Mouse.", Title:="SPECIFY RANGE", Type:=8)
    txtSheetD.Text = oWBD.ActiveSheet.Name
    Set oWSD = oWBD.ActiveSheet
    txtDataD.Text = rRangeD.Address
End Sub

Private Sub Form_Load()
    Set oExcel = New Excel.Application
    oExcel.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
         Set oWSL = Nothing
         Set oWSD = Nothing
    On Error GoTo ss

        If Not oWBL Is Nothing Then oWBL.Close False
        If Not oWBD Is Nothing Then oWBD.Close False
        Set oWBL = Nothing
        Set oWBD = Nothing
        oExcel.Quit
        Set oExcel = Nothing
ss:         Exit Sub
End Sub
