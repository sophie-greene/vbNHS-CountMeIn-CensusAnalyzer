VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmSelect 
   Caption         =   "Choose a file"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save && Exit"
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "A&pply"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame fraCont 
      Height          =   3375
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   6375
      Begin VB.CommandButton cmdCensus 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   16
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtCensus 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdLocal 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   10
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtLocal 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdRegional 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtRegional 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdNational 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   5760
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtNational 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblCensus 
         Caption         =   "Select Census CSV File"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblLocal 
         Caption         =   "Select CMI Local Excel File: "
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblRegional 
         Caption         =   "Select CMI Regional Data"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblNational 
         Caption         =   "Select CMI National Excel File: "
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
   End
   Begin MSComctlLib.TabStrip tbStrp 
      Height          =   4095
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    End If

End Sub

Private Sub cmdApply_Click()
 SaveFiles
End Sub

Private Sub cmdCancel_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdCensus_Click(Index As Integer)
    '-------------------------------------
'The Open command button is clicked.
'-------------------------------------

'If an error occurs goto ErrorH(below)...
On Error GoTo ErrorH

'This setups the dialog box to only
'allow Icons to be loaded...
CommonDialog.Filter = "Comma seperate File|*.csv"

'The Open dialog box is shown...
CommonDialog.ShowOpen
txtCensus(Index).Text = CommonDialog.FileName
'If an error occurs(usually when Cancel is clicked)
ErrorH:
Common_Error 'Call the Common_Error sub-routine...
Exit Sub         'Exit the sub(if an error)...
End Sub

Private Sub cmdLocal_Click(Index As Integer)
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
txtLocal(Index).Text = CommonDialog.FileName
'If an error occurs(usually when Cancel is clicked)
ErrorH:
Common_Error 'Call the Common_Error sub-routine...
Exit Sub         'Exit the sub(if an error)...
End Sub

Private Sub cmdNational_Click(Index As Integer)
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
    txtNational(Index).Text = CommonDialog.FileName
    'If an error occurs(usually when Cancel is clicked)
ErrorH:
    Common_Error 'Call the Common_Error sub-routine...
    Exit Sub         'Exit the sub(if an error)...

End Sub

Private Sub cmdRegional_Click(Index As Integer)
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
    txtRegional(Index).Text = CommonDialog.FileName
    'If an error occurs(usually when Cancel is clicked)
ErrorH:
    Common_Error 'Call the Common_Error sub-routine...
    Exit Sub         'Exit the sub(if an error)...
End Sub

Private Sub cmdSave_Click()
    SaveFiles
    Unload Me
End Sub

Private Sub Form_Load()
    tbStrp.Tabs.Remove (1)
    formatControls
    tbStrp.Tabs(Year(Date) - 2005 + 1).Selected = True
    loadFiles
End Sub


Private Sub tbStrp_Click()
    showframe (tbStrp.SelectedItem - 2005)
End Sub

Private Sub showframe(intX As Integer)
    Dim intC As Integer
    
    For intC = fraCont.lbound To fraCont.UBound
        If intC = intX Then
            fraCont(intC).Visible = True
        Else
            fraCont(intC).Visible = False
        End If
    Next intC
End Sub

Private Sub formatControls()
    Dim i As Integer
    Me.Width = 3 * Screen.Width / 4
    Me.Height = 3 * Screen.Height / 4
    Me.Left = Screen.Width / 8
    Me.Top = Screen.Width / 16
    tbStrp.Left = (Me.Width / 32)
    tbStrp.Width = Me.Width - 2 * tbStrp.Left
    tbStrp.Top = (Me.Height / 32)
    tbStrp.Height = 0.8 * Me.Height - 2 * tbStrp.Top
    tbStrp.TabFixedHeight = 200
    tbStrp.SelectedItem = Right((Year(Date) - 2005), 2)
    fraCont.Item(0).Width = 0.95 * tbStrp.Width
    fraCont.Item(0).Height = (0.95 * tbStrp.Height) - 2 * tbStrp.TabFixedHeight
    fraCont.Item(0).Left = tbStrp.Left + 0.025 * tbStrp.Width
    fraCont.Item(0).Top = 2 * tbStrp.TabFixedHeight + tbStrp.Top
    
    cmdApply.Left = (Me.Width / 2) - cmdCancel.Width / 2
    cmdSave.Left = cmdApply.Left - 1.5 * cmdApply.Width
    cmdCancel.Left = cmdApply.Left + 1.5 * cmdApply.Width
    
    cmdSave.Top = fraCont(0).Top + tbStrp.Height
    cmdCancel.Top = cmdSave.Top
    cmdApply.Top = cmdSave.Top
    
    lblNational(0).Width = 2300
    lblNational(0).Height = 350
    lblNational(0).Left = 100
    lblNational(0).Top = 350
    
    txtNational(0).Width = 0.92 * fraCont(0).Width - lblNational(0).Width
    txtNational(0).Height = 350
    txtNational(0).Left = lblNational(0).Left + lblNational(0).Width
    txtNational(0).Top = 350
    
    cmdNational(0).Width = 0.95 * fraCont(0).Width - txtNational(0).Width - lblNational(0).Width
    cmdNational(0).Height = 350
    cmdNational(0).Left = txtNational(0).Left + txtNational(0).Width + cmdNational(0).Width
    cmdNational(0).Top = 350
    
    lblRegional(0).Width = lblNational(0).Width
    lblRegional(0).Height = lblNational(0).Height
    lblRegional(0).Left = lblNational(0).Left
    lblRegional(0).Top = 3 * lblNational(0).Top
    
    txtRegional(0).Width = txtNational(0).Width
    txtRegional(0).Height = txtNational(0).Height
    txtRegional(0).Left = txtNational(0).Left
    txtRegional(0).Top = 3 * txtNational(0).Top
    
    cmdRegional(0).Width = cmdNational(0).Width
    cmdRegional(0).Height = cmdNational(0).Height
    cmdRegional(0).Left = cmdNational(0).Left
    cmdRegional(0).Top = 3 * cmdNational(0).Top
    
    lblLocal(0).Width = lblNational(0).Width
    lblLocal(0).Height = lblNational(0).Height
    lblLocal(0).Left = lblNational(0).Left
    lblLocal(0).Top = 5 * lblNational(0).Top
    
    txtLocal(0).Width = txtNational(0).Width
    txtLocal(0).Height = txtNational(0).Height
    txtLocal(0).Left = txtNational(0).Left
    txtLocal(0).Top = 5 * txtNational(0).Top
    
    cmdLocal(0).Width = cmdNational(0).Width
    cmdLocal(0).Height = cmdNational(0).Height
    cmdLocal(0).Left = cmdNational(0).Left
    cmdLocal(0).Top = 5 * cmdNational(0).Top
    
    lblCensus(0).Width = lblNational(0).Width
    lblCensus(0).Height = lblNational(0).Height
    lblCensus(0).Left = lblNational(0).Left
    lblCensus(0).Top = 7 * lblNational(0).Top
    
    txtCensus(0).Width = txtNational(0).Width
    txtCensus(0).Height = txtNational(0).Height
    txtCensus(0).Left = txtNational(0).Left
    txtCensus(0).Top = 7 * txtNational(0).Top
    
    cmdCensus(0).Width = cmdNational(0).Width
    cmdCensus(0).Height = cmdNational(0).Height
    cmdCensus(0).Left = cmdNational(0).Left
    cmdCensus(0).Top = 7 * cmdNational(0).Top
    
    For i = 1 To Year(Date) - 2005 + 1
        tbStrp.Tabs.Add i, , (2005 + i - 1)
        If i <> 0 Then
            
            Load fraCont(i)
            fraCont(i).Visible = True
            fraCont(i - 1).ZOrder
            
            Load lblNational(i)
            With lblNational(i)
                Set .Container = fraCont(i)
                .Move lblNational(0).Left, lblNational(0).Top
                .Visible = True
            End With
            
            Load txtNational(i)
            With txtNational(i)
                Set .Container = fraCont(i)
                .Move txtNational(0).Left, txtNational(0).Top
                .Visible = True
            End With
            Load cmdNational(i)
            With cmdNational(i)
                Set .Container = fraCont(i)
                .Move cmdNational(0).Left, cmdNational(0).Top
                .Visible = True
            End With
            
            Load lblRegional(i)
            With lblRegional(i)
                Set .Container = fraCont(i)
                .Move lblRegional(0).Left, lblRegional(0).Top
                .Visible = True
            End With
            
            Load txtRegional(i)
            With txtRegional(i)
                Set .Container = fraCont(i)
                .Move txtRegional(0).Left, txtRegional(0).Top
                .Visible = True
            End With
            Load cmdRegional(i)
            With cmdRegional(i)
                Set .Container = fraCont(i)
                .Move cmdRegional(0).Left, cmdRegional(0).Top
                .Visible = True
            End With
            
            Load lblLocal(i)
            With lblLocal(i)
                Set .Container = fraCont(i)
                .Move lblLocal(0).Left, lblLocal(0).Top
                .Visible = True
            End With
            
            Load txtLocal(i)
            With txtLocal(i)
                Set .Container = fraCont(i)
                .Move txtLocal(0).Left, txtLocal(0).Top
                .Visible = True
            End With
            Load cmdLocal(i)
            With cmdLocal(i)
                Set .Container = fraCont(i)
                .Move cmdLocal(0).Left, cmdLocal(0).Top
                .Visible = True
            End With
            
            Load lblCensus(i)
            With lblCensus(i)
                Set .Container = fraCont(i)
                .Move lblCensus(0).Left, lblCensus(0).Top
                .Visible = True
            End With
            
            Load txtCensus(i)
            With txtCensus(i)
                Set .Container = fraCont(i)
                .Move txtCensus(0).Left, txtCensus(0).Top
                .Visible = True
            End With
            Load cmdCensus(i)
            With cmdCensus(i)
                Set .Container = fraCont(i)
                .Move cmdCensus(0).Left, cmdCensus(0).Top
                .Visible = True
            End With
        End If

           
    Next i
End Sub

Private Sub loadFiles()

    Dim i As Integer
    Set oExcel = New Excel.Application
    'oExcel.Visible = True ' <-- *** Optional ***
    
    Dim oRng1 As Object
    

    Set oWB = oExcel.Workbooks.Open(App.Path & "\files.xls")
    Set oWS = oWB.Worksheets("Sheet1")
    ' A3 to C16 ---> CMI
    For i = 3 To Right$(Year(Date) - 2005, 2) + 3
        txtNational(i - 3).Text = oWS.Cells(i, 2).Value
        txtRegional(i - 3).Text = oWS.Cells(i, 3).Value
        txtLocal(i - 3).Text = oWS.Cells(i, 4).Value
        txtCensus(i - 3).Text = oWS.Cells(i, 5).Value
    Next i
Cleanup:
    Set oWS = Nothing
    If Not oWB Is Nothing Then oWB.Close False
    Set oWB = Nothing
    oExcel.Quit
    Set oExcel = Nothing
End Sub

Private Sub SaveFiles()
  
    Dim i As Integer
    Set oExcel = New Excel.Application
    'oExcel.Visible = True ' <-- *** Optional ***
    
    Dim oRng1 As Object
    

    Set oWB = oExcel.Workbooks.Open(App.Path & "\files.xls", , False)
    Set oWS = oWB.Worksheets("Sheet1")
    ' A3 to C16 ---> CMI
    For i = 3 To Right$(Year(Date) - 2005, 2) + 3
        oWS.Cells(i, 2).Value = txtNational(i - 3).Text
        oWS.Cells(i, 3).Value = txtRegional(i - 3).Text
        oWS.Cells(i, 4).Value = txtLocal(i - 3).Text
        oWS.Cells(i, 5).Value = txtCensus(i - 3).Text
    Next i
    oWB.Save
    
Cleanup:
    Set oWS = Nothing
    If Not oWB Is Nothing Then oWB.Close True
    Set oWB = Nothing
    oExcel.Quit
    Set oExcel = Nothing
End Sub

