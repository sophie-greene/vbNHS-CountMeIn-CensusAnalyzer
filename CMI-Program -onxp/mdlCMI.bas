Attribute VB_Name = "mdlCMI"
Option Explicit
Public oExcel As Excel.Application
Public oWB As Excel.Workbook
Public oWS As Excel.Worksheet
Public oRange As Excel.Range
Public oChart As Excel.Chart
Public oWB1 As Excel.Workbook
Public oWS1 As Excel.Worksheet


'hold template data
Public oWBTemp As Excel.Workbook
Public oWSTemp As Excel.Worksheet

Public intWidth As Double
Public intLeft As Double
Public varChartType As Variant
Public files() As String
Public wsName() As String
Public intColCount() As Integer
Public lookup() As Integer


Sub readSheetNames(strFile As String)
    Dim i As Integer
    Dim owbT As Object
    Dim owsT As Object
    Set owbT = oExcel.Workbooks.Open(strFile)
    ReDim wsName(1 To owbT.Worksheets.Count)
    ReDim intColCount(1 To owbT.Worksheets.Count)
    For i = 1 To owbT.Worksheets.Count
        wsName(i) = owbT.Worksheets.Item(i).Name
        Set owsT = owbT.Worksheets.Item(wsName(i))
        intColCount(i) = owsT.Range("b1").Value
    Next i
Cleanup:
    Set owsT = Nothing
    If Not owbT Is Nothing Then owbT.Close
    Set owbT = Nothing
End Sub

Sub ReadFileNames()
'the raw data files
'Names are stored in files.xls wich is attached to the application
'a3-a& year(date)-2005 +3 -years
'b3-b& year(date)-2005 +3 -National cmi
'c3-c& year(date)-2005 +3 -Regional cmi
'd3-d& year(date)-2005 +3 -local cmi
'e3-e& year(date)-2005 +3 -census
'create a new xls file to arrange data in one place to be able to plot it

    
    Dim i As Integer
    Dim j As Integer


    Set oWB = oExcel.Workbooks.Open(App.Path & "\files")
    
    Set oWS = oWB.Worksheets("Sheet1")
    Set oRange = oWS.Range("a3:e" & (Year(Date) - 2005 + 3))

    ReDim files((Year(Date) - 2005 + 3), 5)
    For i = 3 To (Year(Date) - 2005 + 3)
        For j = 1 To 5
            files(i, j) = oWS.Cells(i, j).Value
            oWS.Cells(i + 10, j + 10).Value = "( " & i & "," & j & ") " & files(i, j)
        Next j
    Next i

Cleanup:
    Set oRange = Nothing
    Set oWS = Nothing
    If Not oWB Is Nothing Then oWB.Close True
    Set oWB = Nothing

End Sub

Sub generateData()
'Years
'( 3,1) 2005
'( 4,1) 2006
'( 5,1) 2007
'( 6,1) 2008
'( 7,1) 2009
'National
'( 3,2) C:\Users\User\Documents\CMI\Count me in project\Count Me in\National\Count me in census 2005 England results.xls
'( 4,2) C:\Users\User\Documents\CMI\Count me in project\Count Me in\National\Count me in census 2006 Mental health results tables ethnicity - England.xls
'( 5,2) C:\Users\User\Documents\CMI\Count me in project\Count Me in\National\Count me in census 2007 Mental health results tables ethnicity - England.xls
'( 6,2) C:\Users\User\Documents\CMI\Count me in project\Count Me in\National\MH_xtab_England_2008.xls
'( 7,2)
'Regional
'( 3,3) C:\Users\User\Documents\CMI\Count me in project\Count Me in\Regional\Yorkshire_and_Humberside2005.xls
'( 4,3) C:\Users\User\Documents\CMI\Count me in project\Count Me in\Regional\Count me in census 2006 Mental health results tables ethnicity - Yorkshire and The Humber SH.xls
'( 5,3) C:\Users\User\Documents\CMI\Count me in project\Count Me in\Regional\Count me in census 2007 Mental health results tables ethnicity - Yorkshire and The Humber SHA.xls
'( 6,3) C:\Users\User\Documents\CMI\Count me in project\Count Me in\Regional\Count me in census 2008 Mental health results tables ethnicity - Yorkshire and The Humber SHA.xls
'( 7,3)
'Local
'( 3,4)
'( 4,4)
'( 5,4)
'( 6,4)
'( 7,4)
'census
'( 3,5) C:\Users\User\Documents\CMI\Count me in project\dataset.csv
'( 4,5) C:\Users\User\Documents\CMI\Count me in project\dataset.csv
'( 5,5) C:\Users\User\Documents\CMI\Count me in project\dataset.csv
'( 6,5) C:\Users\User\Documents\CMI\Count me in project\dataset.csv
'( 7,5) C:\Users\User\Documents\CMI\Count me in project\dataset.csv



    Dim i As Integer
    Dim j As Integer

    Set oExcel = New Excel.Application
    oExcel.Visible = True
    
    readSheetNames (App.Path & "\agregateData")
    ReadFileNames
    copyFileData
Cleanup:
    oExcel.Quit
    Set oExcel = Nothing
    
    
End Sub
Sub copyFileData()
    
    Dim i As Integer
    Dim j As Integer
    Dim x As Integer
    Dim y As Integer
    Dim intX As Integer
    Dim intY As Integer
    
    Dim intCol As Integer
    

    If CheckPath(App.Path & "\analysis.xls") Then Kill App.Path & "\analysis.xls"
    If CheckPath(App.Path & "\analysis.xlsx") Then Kill App.Path & "\analysis.xlsx"

    Set oWB = oExcel.Workbooks.Open(App.Path & "\agregateData.xls")
    For i = 2 To UBound(wsName)
    
        Set oWS = oWB.Worksheets.Item(i)
        Set oWSTemp = oWBTemp.Worksheets.Item(i)
       
        
        For j = 0 To Year(Date) - 2005
        
            oWS.Cells(2 + 30 * j, 3).Value = files(3 + j, 1)
            oWS.Cells(3 + 30 * j, 3).Value = "National-CMI"
            oWS.Cells(3 + 30 * j, 3 + intColCount(i)).Value = "Regional-CMI"
            oWS.Cells(3 + 30 * j, 3 + intColCount(i) * 2).Value = "Local-CMI"
            oWS.Cells.Columns.AutoFit
            
            'National
            oWS.Range(oWS.Cells(4, 1), oWS.Cells(30, 2 + intColCount(i))).Copy _
            oWS.Range(oWS.Cells(4 + (30 * j), 1), oWS.Cells(30 * (j + 1), 2 + intColCount(i)))

            'Regional
            oWS.Range(oWS.Cells(4, 3), oWS.Cells(30, 2 + intColCount(i))).Copy _
            oWS.Range(oWS.Cells(4 + 30 * j, 3 + intColCount(i)), oWS.Cells(4 + 30 * j, 2 + intColCount(i) * 2))
   
            'local
            oWS.Range(oWS.Cells(4, 3), oWS.Cells(30, 2 + intColCount(i))).Copy _
            oWS.Range(oWS.Cells(4 + 30 * j, 3 + intColCount(i) * 2), oWS.Cells(4 + 30 * j, 3 + 2 * intColCount(i)))
    
        
      Next j
   Next i
   

          Set oWSTemp = Nothing
        If Not oWBTemp Is Nothing Then oWBTemp.Close False
        Set oWBTemp = Nothing
        
    For i = 3 To Year(Date) - 2005 + 3
        'For j = 2 To 4
          'National
            If files(i, 2) <> "" Then
                Set oWBTemp = oExcel.Workbooks.Open(files(i, 2))
                For j = 2 To oWBTemp.Worksheets.Count
                     Set oWSTemp = oWBTemp.Worksheets.Item(j)
                    'check if worksheet exists
                    On Error Resume Next
                    Set oWS = oWB.Worksheets(oWBTemp.Worksheets.Item(j).Name)
                    
                    If oWS Is Nothing Then
                

                    Else
                        

    For intX = 1 To oWS.Range(oWS.Cells(6 + (30 * (i - 3)), 3), oWS.Cells(30 * ((i - 2)), 2 + intColCount(j))).Rows.Count
        For intY = 1 To oWS.Range(oWS.Cells(6 + (30 * (i - 3)), 3), oWS.Cells(30 * ((i - 2)), 2 + intColCount(j))).Columns.Count
            oWS.Range(oWS.Cells(6 + (30 * (i - 3)), 3), oWS.Cells(30 * ((i - 2)), 2 + intColCount(j))).Cells(intX, intY).Value = oWSTemp.Range(oWSTemp.Cells(6, 3), oWSTemp.Cells(36, 2 + intColCount(j))).Cells(intX, intY).Value
        Next intY
    Next intX
                      
                    End If
                Next j
            End If
            Set oWSTemp = Nothing
                    If Not oWBTemp Is Nothing Then oWBTemp.Close False
                    Set oWBTemp = Nothing
    Next i
          
           
         'For i = 3 To Year(Date) - 2005 + 3
            'Regional
            'If files(i, 3) <> "" Then
                'Set oWBTemp = oExcel.Workbooks.Open(files(i, 3))
                'For j = 2 To oWBTemp.Worksheets.Count
                    'Set oWSTemp = oWBTemp.Worksheets.Item(j)
                    'check if worksheet exists
                    'On Error Resume Next
                    'Set oWS = oWB.Worksheets(oWBTemp.Worksheets.Item(j).Name)
                    
                    'If oWS Is Nothing Then
                

                    'Else
                        'oWSTemp.Range(oWSTemp.Cells(4, 3), oWSTemp.Cells(30, 2 + intColCount(i))).Copy _
                        'oWS.Range(oWS.Cells(4 + 30 * (i - 3), 3 + intColCount(j)), oWS.Cells(4 + 30 * (i - 2), 2 + intColCount(j) * 2))
     
                    'End If
 
                'Next j
            'Set oWSTemp = Nothing
            'If Not oWBTemp Is Nothing Then oWBTemp.Close False
            'Set oWBTemp = Nothing
            'End If
            
        'Next j
    'Next i
   
    oWB.SaveAs App.Path & "\analysis"
   
    Set oWS = Nothing
    If Not oWB Is Nothing Then oWB.Close
    Set oWB = Nothing
End Sub


Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function
                     

