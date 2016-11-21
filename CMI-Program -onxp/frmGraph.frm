VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmGraph 
   Caption         =   "Graphical and Tabular representations"
   ClientHeight    =   8100
   ClientLeft      =   -1185
   ClientTop       =   795
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk 
      Caption         =   "Show Census Baseline"
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H0000C000&
      Caption         =   "&Generate Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "based on your choises, generates excel chart and display it"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9600
      TabIndex        =   18
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Export to Excel"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Frame fraMain 
      Height          =   6615
      Left            =   360
      TabIndex        =   19
      Top             =   -120
      Width           =   13335
      Begin VB.Frame fraMCat 
         BorderStyle     =   0  'None
         Caption         =   "Please choose the categories you want to compare"
         Height          =   3855
         Left            =   1080
         TabIndex        =   12
         Top             =   2040
         Width           =   4455
         Begin VB.ListBox lstMain 
            Height          =   2985
            ItemData        =   "frmGraph.frx":0000
            Left            =   840
            List            =   "frmGraph.frx":0002
            MultiSelect     =   1  'Simple
            TabIndex        =   14
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label Label1 
            Caption         =   "Please choose the categories you want to compare"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame fraSource 
         Caption         =   "Please Choose Data Source"
         Height          =   1335
         Left            =   8640
         TabIndex        =   5
         Top             =   240
         Width           =   3735
         Begin VB.CheckBox chckSource 
            Caption         =   "Local"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox chckSource 
            Caption         =   "Regional"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chckSource 
            Caption         =   "National"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraSec 
         BorderStyle     =   0  'None
         Caption         =   "Please choose the categories you want to compare"
         Height          =   3855
         Left            =   720
         TabIndex        =   9
         Top             =   1920
         Width           =   4455
         Begin VB.ListBox lstSecondry 
            Height          =   2985
            Left            =   480
            MultiSelect     =   1  'Simple
            TabIndex        =   11
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblList 
            Caption         =   "Please choose the categories you want to compare"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame fraGraph 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   4920
         TabIndex        =   21
         Top             =   1800
         Width           =   4335
         Begin VB.OLE oleGraph 
            BackStyle       =   0  'Transparent
            BorderStyle     =   0  'None
            Class           =   "Excel.Chart.8"
            Height          =   3495
            Left            =   360
            SizeMode        =   1  'Stretch
            TabIndex        =   22
            Top             =   600
            Width           =   3735
         End
      End
      Begin VB.Frame fraChart 
         Caption         =   "Choose Comparison Subject"
         Height          =   975
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   4335
         Begin VB.ComboBox cmbField 
            Height          =   315
            ItemData        =   "frmGraph.frx":0004
            Left            =   2400
            List            =   "frmGraph.frx":0006
            TabIndex        =   4
            Text            =   "Choose comparison field"
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cmbGraph 
            Height          =   315
            ItemData        =   "frmGraph.frx":0008
            Left            =   480
            List            =   "frmGraph.frx":000A
            TabIndex        =   3
            Text            =   "Choose comparison area"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblChart 
            Height          =   495
            Left            =   480
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fraCat 
         Caption         =   "Please Choose Categories"
         Height          =   1455
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   3735
         Begin VB.ComboBox cmbCat 
            Height          =   315
            ItemData        =   "frmGraph.frx":000C
            Left            =   120
            List            =   "frmGraph.frx":0019
            TabIndex        =   1
            Text            =   "Choose Category"
            Top             =   720
            Width           =   3375
         End
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   90
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image imgLogo 
      Height          =   1860
      Left            =   0
      Picture         =   "frmGraph.frx":00EF
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1635
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rGraph As Excel.Range
Dim xlSeries As Excel.Series
Dim oWBGraphs As Excel.Workbook
Dim oWSGraphs As Excel.Worksheet
Dim noPrimary(0 To 4) As Integer
Dim strPrimary(0 To 5) As String
Dim strSecondary(0 To 16) As String
Const noWhite As Integer = 3
Const noMixed As Integer = 4
Const noAsian As Integer = 4
Const noBlack As Integer = 3
Const noOther As Integer = 2
Const noBME As Integer = 15
Const noTotal As Integer = 18

Dim boolBaseline As Boolean
Dim intCat As Integer '0 main,1 primary, 2 secondary
Dim arrEthnicities(5) As Integer
Dim strSheet As String
Dim intSource(3) As Integer
Dim intFields As Integer
Dim intStart As Integer
Dim intEnd As Integer
Dim intField As Integer
Dim strFormat As String

'strShteet is the source sheet selected in cmbGraph
'intSource() array of the check boxes indeces to indicate which source (National,Regional,Local)
'intField indicate the col in the sheet for example FOR "Arabic" IN LANGUAGES 0


Private Sub chckSource_Click(Index As Integer)
    Dim intX
    For intX = 0 To 2
        If chckSource(intX).Enabled = True Then
        intSource(intX) = chckSource(intX).Value
        End If
    Next intX
    oleGraph.Visible = False
End Sub
Function getGroups() As Integer
    Dim i As Integer
    getGroups = 0
    For i = 0 To 2
        If intSource(i) = 1 Then
            getGroups = getGroups + 1
        End If
    Next i

End Function

Private Sub chk_Click()
    If chk.Value = 1 Then
        boolBaseline = True
    Else
        boolBaseline = False
    End If
End Sub

Private Sub cmbCat_click()
    'main
    If cmbCat.Text = "Main Categories(White, BME, Not stated)" Then
        fraMCat.Visible = False
        fraMCat.Enabled = False
        fraSec.Visible = False
        fraSec.Enabled = False
        intWidth = fraGraph.Width
        intLeft = fraGraph.Left
        fraGraph.Width = fraMain.Width - 2 * fraCat.Left
        fraGraph.Left = fraCat.Left
        fraGraph.Top = fraMCat.Top
        intCat = 0
    ElseIf cmbCat.Text = "Primary Categories(White, Mixed, Asian, Black, other, Not stated)" Then
        fraMCat.Visible = True
        fraMCat.Enabled = True
        fraSec.Visible = False
        fraSec.Enabled = False
        fraGraph.Width = intWidth
        fraGraph.Left = intLeft
        oleGraph.Left = 0
        oleGraph.Top = 0
        oleGraph.Width = fraGraph.Width
        oleGraph.Height = fraGraph.Height
        intCat = 1
    ElseIf cmbCat.Text = "Secondary Categories(White British, White Irish, White Other, Mixed While and Black Caribbean,....)" Then
        fraMCat.Visible = False
        fraMCat.Enabled = False
        fraSec.Visible = True
        fraSec.Enabled = True
        fraGraph.Width = intWidth
        fraGraph.Left = intLeft
        intCat = 2
    End If
        oleGraph.Left = 0
        oleGraph.Top = 0
        oleGraph.Width = fraGraph.Width
        oleGraph.Height = fraGraph.Height
        oleGraph.Visible = False
'oleGraph.Enabled = False
On Error Resume Next
End Sub

Private Sub cmbField_click()
cmbCat.Clear
    If oWS.Cells(1, 8).Value Then
        cmbCat.AddItem "Main Categories(White, BME, Not stated)"
        cmbCat.AddItem "Primary Categories(White, Mixed, Asian, Black, other, Not stated)"
    End If
    cmbCat.AddItem "Secondary Categories(White British, White Irish, White Other, Mixed While and Black Caribbean,....)"
    cmbCat.Enabled = True
    intField = cmbField.ListIndex + 1
    oleGraph.Visible = False
End Sub

Private Sub cmbGraph_click()
 strSheet = oWB.Worksheets.Item(cmbGraph.List(cmbGraph.ListIndex)).Name
    cmbField.Enabled = True
    cmbField.Clear
    cmbCat.Clear
Dim intX As Integer
    Set oWS = oWB.Worksheets.Item(cmbGraph.Text)
    intFields = oWS.Cells(1, 1).Value
    intStart = oWS.Cells(1, 2).Value
    intEnd = oWS.Cells(1, 3).Value
    strFormat = oWS.Cells(1, 4).Value
    For intX = 0 To intFields - 1
        cmbField.AddItem oWS.Cells(intStart - 1, intX + 3).Value
    Next intX
    cmbField.Text = "Choose area of Comparison"

    cmbCat.Text = "Choose Category"
    If oWS.Cells(1, 5).Value = "std" Then
        chckSource(0).Enabled = True
        chckSource(1).Enabled = False
        intSource(1) = 0
        chckSource(2).Enabled = False
        intSource(2) = 0
    Else
        chckSource(0).Enabled = True
        chckSource(1).Enabled = True
        chckSource(2).Enabled = True
    End If
    If oWS.Cells(1, 4).Value = "#" Then
    chk.Enabled = False
    Else
    chk.Enabled = True
    End If
    oleGraph.Visible = False
    lblInfo.Caption = oWS.Cells(1, 6).Value & "---> " & oWS.Cells(1, 7).Value
End Sub

Private Sub cmdExcel_Click()
        ' Sets the Dialog Title to Save File
    CommonDialog1.DialogTitle = "Save Excel File"
    
    ' Sets the File List box to Text File and All Files
    CommonDialog1.Filter = "Excel Files 97-2003 (*.xls)|*.xls|Excel Files 2007(*.xlsx)|*.xlsx|AllFiles (*.*)|*.*"
    
    ' Set the default files type to Text File
    CommonDialog1.InitDir = Environ$("USERPROFILE") & "\Desktop\"
    
    ' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt _
    + cdlOFNPathMustExist
    
    ' Set dialog box so an error occurs if the dialogbox is cancelled
    CommonDialog1.CancelError = True
    
    ' Enables error handling to catch cancel error
    On Error Resume Next
    ' display the dialog box
    CommonDialog1.ShowSave
    If Err Then
        ' This code runs if the dialog was cancelled
        MsgBox "Dialog Cancelled"
        Exit Sub
    End If
    FileCopy App.Path & "\agregateData.xls", CommonDialog1.FileName
    Dim e As Excel.Application
    Set e = New Excel.Application
     e.Workbooks.Open (CommonDialog1.FileName)
    ' e.Visible = True
End Sub

Private Sub cmdExit_Click()
    'frmMain.Show
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    If Not isEssential Then
        MsgBox "please select the  attributes highlighted in red"

        GoTo endof
    End If
    Set oWS = oWB.Worksheets.Item(strSheet)
    'oExcel.Visible = True
    If Not oWBGraphs Is Nothing Then oWBGraphs.Close False
    Set oWBGraphs = Nothing
    plotGraph

    If CheckPath(App.Path & "\res.xls") Then Kill App.Path & "\res.xls"
    If CheckPath(App.Path & "\res.xlsx") Then Kill App.Path & "\res.xlsx"
    oWBGraphs.SaveAs App.Path & "\res.xls"
    oleGraph.Visible = True
    oleGraph.Class = "Excel.Chart8.0"

    oleGraph.CreateLink oWBGraphs.Path & "\" & oWBGraphs.Name, "chart1"

    oleGraph.Update

endof:

End Sub

Private Sub Form_LinkError(LinkErr As Integer)
    MsgBox "an error has occured. please restart the program and try again"
    Resume Next
End Sub

Private Sub Form_Load()
    Set oExcel = New Excel.Application
    'oExcel.Visible = True
    Set oWB = oExcel.Workbooks.Open(App.Path & "\agregateData.xls", False)
    formatControls
    loadLists
    getComparison
    intWidth = fraGraph.Width
    intLeft = fraGraph.Left
    initialiseGlobals
    cmbCat.Enabled = False
    cmbField.Enabled = False
    fraMCat.Visible = False
    fraSec.Visible = False
End Sub


'REGIONAL STARTS AT 2+LENGTH
'LOCAL STARTS AS 2+2*LENGTH

'fill the item to be compared list using the fields available
Sub getComparison()
    Dim intX As Integer

    For intX = 1 To oWB.Worksheets.count - 3
        cmbGraph.AddItem oWB.Worksheets(intX).Name
    Next intX
    cmbGraph.Text = "choose comparison area"
    'Set oWS = oWB.Worksheets.Item(cmbGraph.Text)
End Sub

'**********************************Format Control ****************************************************
Private Sub formatControls()
    Dim i As Integer
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.Left = Screen.Width
    Me.Top = Screen.Width
    
    fraSource.Height = 5 * chckSource(0).Height
    fraSource.Width = (fraMain.Width - 1.5 * fraMain.Left) / 4
    fraMain.Width = 0.9 * Me.Width
    fraMain.Height = 0.78 * Me.Height
    fraMain.Top = (Me.Height - fraMain.Height) / 6
    fraMain.Left = (Me.Width - fraMain.Width) / 2
    
    fraCat.Width = (fraMain.Width - 1.5 * fraMain.Left - fraSource.Width) / 2
    fraCat.Left = fraMain.Left / 3
    fraCat.Height = fraSource.Height
    fraCat.Top = fraMain.Top
    
    fraChart.Width = fraCat.Width
    fraChart.Left = 2 * fraCat.Left + fraCat.Width
    fraChart.Height = fraCat.Height
    fraChart.Top = fraMain.Top
    
    fraSource.Left = fraCat.Left + fraCat.Width + fraChart.Left
    fraSource.Top = fraChart.Top

    cmbGraph.Width = 0.9 * fraChart.Width
    cmbGraph.Left = fraCat.Left
    cmbGraph.Top = fraCat.Top
    
    cmbField.Width = 0.9 * fraCat.Width
    cmbField.Left = fraCat.Left
    cmbField.Top = fraCat.Top + cmbField.Height
    
    cmbCat.Width = 0.95 * fraChart.Width
    cmbCat.Left = cmbGraph.Left / 2
    cmbCat.Top = (fraChart.Height - cmbCat.Height) / 2
    

    
    fraMCat.Top = fraCat.Height + fraCat.Top
    fraMCat.Height = (fraMain.Height - fraMCat.Top - 2 * fraCat.Top)
    fraMCat.Width = 3 * fraCat.Width / 4
    fraMCat.Left = fraCat.Left
    
    fraSec.Left = fraMCat.Left
    fraSec.Width = fraMCat.Width
    fraSec.Top = fraMCat.Top
    fraSec.Height = fraMCat.Height

    fraGraph.Left = fraMCat.Left + fraMCat.Width

    fraGraph.Height = fraMCat.Height
    fraGraph.Width = (fraMain.Width - fraMCat.Width) - 2 * fraMCat.Left
    fraGraph.Top = fraMCat.Top
    
    lstMain.Left = fraMCat.Left
    lstMain.Width = fraMCat.Width - 2 * lstMain.Left
    lstMain.Top = 0.1 * fraMCat.Height
    lstMain.Height = fraMCat.Height - 1.5 * lstMain.Top
    lstSecondry.Top = lstMain.Top
    lstSecondry.Height = lstMain.Height
    lstSecondry.Width = lstMain.Width
    lstSecondry.Left = 0
    
    cmdExcel.Height = (Me.Height - (fraMain.Top + fraMain.Height)) / 4
    cmdExcel.Width = 0.15 * Me.Width
    cmdExcel.Left = (Me.Width - 3 * cmdExcel.Width) / 2
    
    cmdExcel.Top = fraMain.Top + fraMain.Height + cmdExcel.Height / 2
    cmdExit.Height = cmdExcel.Height
    cmdExit.Width = cmdExcel.Width
    cmdExit.Left = cmdExcel.Left + 1.5 * cmdExcel.Width
    cmdExit.Top = cmdExcel.Top
    imgLogo.Height = Me.Height - fraMain.Height - 3 * fraMain.Top
    imgLogo.Top = fraMain.Height + fraMain.Top
    imgLogo.Left = 0
    lblList.Top = lstMain.Top - lblList.Height
    
    chk.Height = cmdExit.Height
    chk.Width = cmdExit.Width
    chk.Left = fraMain.Left + (fraMain.Width - cmdGenerate.Width)
    chk.Top = fraMain.Top + (fraMain.Height)
    
    cmdGenerate.Width = 2 * chk.Width / 3
    cmdGenerate.Height = chk.Height
    cmdGenerate.Top = chk.Top + cmdGenerate.Height
    cmdGenerate.Left = chk.Left
        
    lblInfo.Height = fraCat.Height
    lblInfo.Top = fraGraph.Top + fraGraph.Height
    lblInfo.Width = fraMain.Width - fraCat.Left
    lblInfo.Left = fraCat.Left
End Sub

Private Sub loadLists()
    lstSecondry.AddItem "White British"
    lstSecondry.AddItem "White Irish"
    lstSecondry.AddItem "White Other White"
    lstSecondry.AddItem "Mixed White and Black Caribbean"
    lstSecondry.AddItem "Mixed White and Black African"
    lstSecondry.AddItem "Mixed White And Asian"
    lstSecondry.AddItem "Mixed Other Mixed"
    lstSecondry.AddItem "Asian or Asian British Indian"
    lstSecondry.AddItem "Asian or Asian British Pakistani"
    lstSecondry.AddItem "Asian or Asian British Bangladeshi"
    lstSecondry.AddItem "Asian or Asian British Other Asian"
    lstSecondry.AddItem "Black or Black British Caribbean"
    lstSecondry.AddItem "Black or Black British African"
    lstSecondry.AddItem "Black or Black British Other Black"
    lstSecondry.AddItem "Other Ethnic Groups Chinese"
    lstSecondry.AddItem "Other Ethnic Groups Other"


    strSecondary(0) = "White British"
    strSecondary(1) = "White Irish"
    strSecondary(2) = "White Other"
    strSecondary(3) = "Mixed Caribbean"
    strSecondary(4) = "Mixed African"
    strSecondary(5) = "Mixed White And Asian"
    strSecondary(6) = "Mixed Other"
    strSecondary(7) = "Indian"
    strSecondary(8) = "Pakistani"
    strSecondary(9) = "Bangladeshi"
    strSecondary(10) = "Other Asian"
    strSecondary(11) = "Caribbean"
    strSecondary(12) = "African"
    strSecondary(13) = "Other Black"
    strSecondary(14) = "Chinese"
    strSecondary(15) = "Other"


    lstMain.AddItem "White"
    lstMain.AddItem "Mixed"
    lstMain.AddItem "Asian or Asian British"
    lstMain.AddItem "Black or Black British"
    lstMain.AddItem "Other Ethnic Groups"
      
    strPrimary(0) = "White"
    strPrimary(1) = "Mixed"
    strPrimary(2) = "Asian"
    strPrimary(3) = "Black"
    strPrimary(4) = "Other"

    noPrimary(0) = 3 'white
    noPrimary(1) = 4 'mixed
    noPrimary(2) = 4 'asian
    noPrimary(3) = 3 'black
    noPrimary(4) = 2 'other
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cleanup:
       Set oWS = Nothing
       Set oWSGraphs = Nothing
        If Not oWB Is Nothing Then oWB.Close False
        If Not oWBGraphs Is Nothing Then oWBGraphs.Close True
        Set oWB = Nothing
        Set oWBGraphs = Nothing
        oExcel.Quit
        Set oExcel = Nothing
End Sub





Private Sub lstMain_Click()
    Dim i As Integer
    Dim k As Integer
    If lstMain.SelCount = 4 Then
        MsgBox "you are allowed three categories only." & vbCrLf & "If you need to change the selected items click on them again"
        lstMain.Selected(lstMain.ListIndex) = False
        lstMain.ListIndex = -1
    End If
    
    
    'reset
    For i = 0 To UBound(arrEthnicities)
        arrEthnicities(i) = -10
    Next i
    k = 0
    For i = 0 To lstMain.ListCount - 1
        If lstMain.Selected(i) = True Then
            arrEthnicities(k) = i
            k = k + 1
        End If
    Next i
End Sub

Private Sub lstSecondry_Click()
    Dim i As Integer
    Dim k As Integer
    If lstSecondry.SelCount = 6 Then
        MsgBox "you are allowed five categories only." & vbCrLf & "If you need to change the selected items click on them again"
        lstSecondry.Selected(lstSecondry.ListIndex) = False
        lstSecondry.ListIndex = -1
    End If
    
'reset
    For i = 0 To UBound(arrEthnicities)
        arrEthnicities(i) = -10
    Next i
    k = 0
    For i = 0 To lstSecondry.ListCount - 1
        If lstSecondry.Selected(i) = True Then
            arrEthnicities(k) = i
            k = k + 1
        End If
    Next i
End Sub
Function getSel() As Integer
    Dim i As Integer
    getSel = 0
    For i = 0 To UBound(arrEthnicities) - 1
        If arrEthnicities(i) >= 0 Then getSel = getSel + 1
    Next i
    If intCat = 0 Then getSel = 2
End Function

Sub plotGraph()
    Dim j, i, k, ss, m, count As Integer
    Dim dblSum As Double
    Dim arrPrimary(5) As Double
    Dim c As Integer
    Dim noOfCat As Integer
    Dim noOfGroups As Integer
    noOfCat = getSel
    noOfGroups = getGroups
    'data to be copied based on the category
    'create a new sheet to store results
    'source sheet

    Set oWBGraphs = oExcel.Workbooks.Add
    
    Set oWSGraphs = oWBGraphs.Worksheets.Item(1)
    oWSGraphs.Name = strSheet
    'set
    oWSGraphs.Range("A2") = "Years"
    For j = 2005 To Year(Date) - 1
        oWSGraphs.Range("A" & j - 2002).Value = j
        oWSGraphs.Range("A" & j - 2002).NumberFormat = "@"
    Next j
    
    If intCat = 0 Then
        count = 2
        For m = 0 To 2
            If intSource(m) = 1 Then
                oWSGraphs.Cells(2, count).Value = "White British"
                oWSGraphs.Cells(2, count + 1).Value = "BME"
                count = count + 2
            End If
        Next m
    ElseIf intCat = 1 Then
    count = 1
        For m = 0 To 2
            For i = 0 To UBound(arrEthnicities) - 1
                If intSource(m) = 1 Then
                    If arrEthnicities(i) >= 0 Then
                    count = count + 1
                        oWSGraphs.Cells(2, count).Value = strPrimary(arrEthnicities(i))
                    End If
                End If
            Next i
          Next m
    ElseIf intCat = 2 Then
        count = 1
        For m = 0 To 2
            For i = 0 To UBound(arrEthnicities) - 1
                If intSource(m) = 1 Then
                    If arrEthnicities(i) >= 0 Then
                        count = count + 1
                        oWSGraphs.Cells(2, count).Value = strSecondary(arrEthnicities(i))
 
                    End If
                End If
            Next i
      Next m
    End If
    count = 2
    For m = 0 To 2
        For i = 0 To noOfCat - 1
            If intSource(m) = 1 Then
                If m = 0 Then
                    oWSGraphs.Cells(2, count).Value = oWSGraphs.Cells(2, count).Value & " National(CMI)"
                ElseIf m = 1 Then
                    oWSGraphs.Cells(2, count).Value = oWSGraphs.Cells(2, count).Value & " Regional(CMI)"
                ElseIf m = 2 Then
                    oWSGraphs.Cells(2, count).Value = oWSGraphs.Cells(2, count).Value & " Local(CMI)"
                End If
                count = count + 1
            End If
        Next i
    Next m
    For j = 0 To Year(Date) - 2006
        If intCat = 0 Then 'main
            count = 1
            For m = 0 To 2
                If intSource(m) = 1 Then
                    dblSum = 0
                    For k = 0 To noBME - 1
                        If Not oWS.Cells(intStart + 1 + k + (30 * j), intField + 2 + m).Value = "" Then
                        dblSum = dblSum + oWS.Cells(intStart + 1 + k + (30 * j), intField + 2 + m).Value
                        End If
                    Next k
                    oWSGraphs.Cells(j + 3, count + 1).Value = oWS.Cells(intStart + (30 * j), intField + 2 + m).Value
                    oWSGraphs.Cells(j + 3, count + 1).NumberFormat = strFormat
                    If dblSum > 0 Then
                    oWSGraphs.Cells(j + 3, count + 2).Value = dblSum
                    Else
                    oWSGraphs.Cells(j + 3, count + 2).Value = ""
                    End If
                    oWSGraphs.Cells(j + 3, count + 2).NumberFormat = strFormat
                    count = count + 2
                End If
            Next m
        ElseIf intCat = 1 Then 'primary
        'add the primary cat
        count = 1
        For m = 0 To 2
            For k = 0 To UBound(arrPrimary)
                arrPrimary(k) = 0 '0
            Next k
            For i = 0 To UBound(arrEthnicities) - 1
                

                    If intSource(m) = 1 Then
                        If arrEthnicities(i) >= 0 Then
                            c = 0
                            For ss = 0 To arrEthnicities(i) - 1
                                c = c + noPrimary(ss)
                            Next ss
                            
                                For k = 0 To noPrimary(arrEthnicities(i)) - 1
                                    If Not oWS.Cells(intStart + k + c + (30 * j), intField + 2 + m).Value = "" Then
                                    arrPrimary(arrEthnicities(i)) = arrPrimary(arrEthnicities(i)) + _
                                    oWS.Cells(intStart + k + c + (30 * j), intField + 2 + m).Value
                                    End If
                                Next k
                            count = count + 1
                            If arrPrimary(arrEthnicities(i)) > 0 Then
                            oWSGraphs.Cells(j + 3, count).Value = arrPrimary(arrEthnicities(i))
                            Else
                            oWSGraphs.Cells(j + 3, count).Value = ""
                            End If
                            oWSGraphs.Cells(j + 3, count).NumberFormat = strFormat
                    
                        End If
                    End If
                
                Next i
        Next m

        ElseIf intCat = 2 Then 'secondary
            count = 1
            ss = (noOfGroups * noOfCat)
            For m = 0 To 2
                If intSource(m) = 1 Then
                
                    For i = 0 To UBound(arrEthnicities) - 1
                    
                        If arrEthnicities(i) >= 0 Then
                            oWSGraphs.Cells(j + 3, count + 1).Value = oWS.Cells(intStart + arrEthnicities(i) + (30 * j), m + 2 + intField)
                            oWSGraphs.Cells(j + 3, count + 1).NumberFormat = strFormat
                            count = count + 1
                        End If
                    Next i
                    
                End If
            Next m
        End If
        Next j
         ss = (noOfGroups * noOfCat)
        If boolBaseline = True Then
        
            For j = 0 To Year(Date) - 2006
                If intCat = 0 Then
                
                    Set oWS = oWB.Worksheets("censusM")
                    count = 0
                    For m = 0 To 2
                        If intSource(m) = 1 Then
                                If m = 0 Then
                                    oWSGraphs.Cells(2, ss + 2 + count).Value = "White British National(Census)"
                                    oWSGraphs.Cells(2, ss + 3 + count).Value = "BME National(Census)"
                                ElseIf m = 1 Then
                                    oWSGraphs.Cells(2, ss + 2 + count).Value = "White British Regional(Census)"
                                    oWSGraphs.Cells(2, ss + 3 + count).Value = "BME Regional(Census)"
                                ElseIf m = 2 Then
                                    oWSGraphs.Cells(2, ss + 2 + count).Value = "White British Local(Census)"
                                    oWSGraphs.Cells(2, ss + 3 + count).Value = "BME Local(Census)"
                                End If
                                If strFormat = "#" Then
                                    oWSGraphs.Cells(j + 3, ss + 2 + count).Value = oWS.Cells(3, m + 2).Value * oWS.Cells(5, m + 2).Value
                                    oWSGraphs.Cells(j + 3, ss + 3 + count).Value = oWS.Cells(4, m + 2).Value * oWS.Cells(5, m + 2).Value
                                Else
                                    oWSGraphs.Cells(j + 3, ss + 2 + count).Value = oWS.Cells(3, m + 2).Value
                                    oWSGraphs.Cells(j + 3, ss + 3 + count).Value = oWS.Cells(4, m + 2).Value
                                End If
                                oWSGraphs.Cells(j + 3, ss + 3 + count).NumberFormat = strFormat
                                oWSGraphs.Cells(j + 3, ss + 2 + count).NumberFormat = strFormat
                                count = count + 2
                        End If
                    Next m
                ElseIf intCat = 1 Then
                    Set oWS = oWB.Worksheets("censusP")
                    count = 0
                    For m = 0 To 2
                        For i = 0 To UBound(arrEthnicities) - 1
                            If intSource(m) = 1 Then
                                If arrEthnicities(i) >= 0 Then
                                    If m = 0 Then
                                       
                                        
                                        oWSGraphs.Cells(2, ss + 2 + count).Value = strPrimary(arrEthnicities(i)) & " National(Census)"
                                    ElseIf m = 1 Then
                                        oWSGraphs.Cells(2, ss + 2 + count).Value = strPrimary(arrEthnicities(i)) & " Regional(Census)"
                                    ElseIf m = 2 Then
                                        oWSGraphs.Cells(2, ss + 2 + count).Value = strPrimary(arrEthnicities(i)) & " Local(Census)"
                                    End If
                                    If strFormat = "#" Then
                                         oWSGraphs.Cells(j + 3, 2 + ss + count).Value = oWS.Cells(3 + arrEthnicities(i), m + 2).Value * oWS.Cells(8, m + 2).Value
                                    Else
                                    oWSGraphs.Cells(j + 3, 2 + ss + count).Value = oWS.Cells(3 + arrEthnicities(i), m + 2).Value
                                    End If
                                    oWSGraphs.Cells(j + 3, 2 + ss + count).NumberFormat = strFormat
                                    count = count + 1
                                End If
                            End If
                        Next i
                    Next m
                    
                ElseIf intCat = 2 Then
                    Set oWS = oWB.Worksheets("censusS")
                    count = 0
                    For m = 0 To 2
                        For i = 0 To UBound(arrEthnicities) - 1
                            If intSource(m) = 1 Then
                                If arrEthnicities(i) >= 0 Then
                                    If m = 0 Then
                                        oWSGraphs.Cells(2, ss + 2 + count).Value = strSecondary(arrEthnicities(i)) & " National(Census)"
                                    ElseIf m = 1 Then
                                        oWSGraphs.Cells(2, ss + 2 + count).Value = strSecondary(arrEthnicities(i)) & " Regional(Census)"
                                    ElseIf m = 2 Then
                                        oWSGraphs.Cells(2, ss + 2 + count).Value = strSecondary(arrEthnicities(i)) & " Local(Census)"
                                    End If
                                    If strFormat = "#" Then
                                         oWSGraphs.Cells(j + 3, 2 + ss + count).Value = oWS.Cells(3 + arrEthnicities(i), m + 2).Value * oWS.Cells(19, m + 2).Value
                                    Else
                                    oWSGraphs.Cells(j + 3, 2 + ss + count).Value = oWS.Cells(3 + arrEthnicities(i), m + 2).Value
                                    End If
                                    oWSGraphs.Cells(j + 3, 2 + ss + count).NumberFormat = strFormat
                                    count = count + 1
                                End If
                            End If
                        Next i
                    Next m
                End If
            Next j
            If intCat = 0 Then
               ss = ss * 2
            Else
                ss = ss * 2
            End If
            End If
                
        
    oWSGraphs.Cells.Columns.AutoFit
    Set oChart = oWSGraphs.Parent.Charts.Add
    Set rGraph = oWSGraphs.Range("a2").Resize(Year(Date) - 2003, (ss + 1))
    With oChart
       ' .PlotArea.ti
        .SetSourceData Source:=rGraph, PlotBy:=Excel.xlColumns
        .HasTitle = True
      .ChartTitle.Text = strSheet
       .ChartType = Excel.XlChartType.xlLineMarkers
      .Axes(xlValue).HasTitle = True
      .Axes(xlValue).AxisTitle.Text = cmbField.List(intField - 1) & " " & Right(strFormat, 1)
      .Axes(xlPrimary).HasTitle = True
      .Axes(xlPrimary).AxisTitle.Text = "Years"
       '.ChartWizard oWSGraphs.Range(oWSGraphs.Cells(2, 2), oWSGraphs.Cells(2, ss)), , , "years",
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        .HasTitle = True
        .Axes(xlValue, xlPrimary).HasMajorGridlines = False
        For i = 1 To (ss)
            With .SeriesCollection(i).Border
            '.ColorIndex = Int((25) * Rnd + 1)

            .Weight = 3
            .LineStyle = xlContinuous
            
            End With
               
        .SeriesCollection(i).XValues = oWSGraphs.Range("A3", "A6")
                   
        .SeriesCollection(i).Name = oWSGraphs.Cells(2, i + 1).Value
        .SeriesCollection(i).Values = oWSGraphs.Range(oWSGraphs.Cells(3, i + 1), oWSGraphs.Cells(6, i + 1)).Value
        
        Randomize (Timer)
        '.SeriesCollection(i).Points((i Mod (Year(Date) - 2005)) + 1).ApplyDataLabels AutoText:=True, _
        'LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, _
        'ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False
        '.SeriesCollection(i).DataLabels.Position = Rnd
        '.SeriesCollection(i).DataLabels.HorizontalAlignment = xlRight
        '.SeriesCollection(i).DataLabels.VerticalAlignment = xlRight

        '.SeriesCollection(i).DataLabels.Font.Size = 11
        '.SeriesCollection(i).DataLabels.Orientation = 7
        
            Next i
            .SeriesCollection(ss + 1).Delete
            .Location xlLocationAsNewSheet, "Chr"
        End With


End Sub

Sub initialiseGlobals()
    Dim i As Integer
    
    intCat = -10
    intField = -10
    For i = 0 To 5
        arrEthnicities(i) = -10
    Next i
    strSheet = "None"
    For i = 0 To 2
        intSource(i) = -1
    Next i
End Sub
Function isEssential() As Boolean

    If getSel() = 0 Then
        lstMain.ForeColor = vbRed
        lstSecondry.ForeColor = vbRed
    Else
        lstMain.ForeColor = vbBlack
        lstSecondry.ForeColor = vbBlack
    End If
    
    If intCat = -10 Then
        cmbCat.ForeColor = vbRed

    Else
        cmbCat.ForeColor = vbBlack
    End If
    
    If strSheet = "None" Then
        cmbGraph.ForeColor = vbRed

    Else
        cmbGraph.ForeColor = vbBlack
    End If
    
    If strSheet = "None" Then
        cmbGraph.ForeColor = vbRed

    Else
        cmbGraph.ForeColor = vbBlack
    End If
    
    If intField = -10 Then
        cmbField.ForeColor = vbRed

    Else
        cmbField.ForeColor = vbBlack
    End If
    
    If (intSource(0) = -1 And intSource(1) = -1 And intSource(2) = -1) Then
        fraSource.ForeColor = vbRed

    Else
        fraSource.ForeColor = vbBlack
    End If
      If (intSource(0) = 0 And intSource(1) = 0 And intSource(2) = 0) Then
        fraSource.ForeColor = vbRed

    Else
        fraSource.ForeColor = vbBlack
    End If
    isEssential = True
    If getSel() = 0 Or intCat = -1 Or _
    strSheet = "None" _
    Or intField = -10 _
    Or (intSource(0) = -1 And intSource(1) = -1 And intSource(2) = -1) Or intSource(0) = 0 And intSource(1) = 0 And intSource(2) = 0 Then
        
        isEssential = False

    End If
End Function
