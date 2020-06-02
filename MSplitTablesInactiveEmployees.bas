Attribute VB_Name = "MSplitTables4"
Const Target_Folder As String = "C:\Users\I519797\OneDrive - SAP SE\Documents\Work\2020\Week 23\"
Dim wsSource As Worksheet, wsHelper As Worksheet
Dim LastRow As Long, LastColumn As Long
 
Sub SplitDataset()
 
    Dim collectionUniqueList As Collection
    Dim i As Long
 
    Set collectionUniqueList = New Collection
 
    Set wsSource = ThisWorkbook.Worksheets("Sheet2")
    Set wsHelper = ThisWorkbook.Worksheets("Sheet1")
 
 
    wsHelper.Cells.ClearContents
 
    With wsSource
        .AutoFilterMode = False
 
        LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
 
        If .Range("A2").Value = "" Then
            GoTo Cleanup
        End If
 
        Call Init_Unique_List_Collection(collectionUniqueList, LastRow)
 
        Application.DisplayAlerts = False
 
        For i = 1 To collectionUniqueList.Count
                 If InStr(collectionUniqueList.Item(i), "BELUX") > 0 Or InStr(collectionUniqueList.Item(i), "FRANCE") > 0 Or InStr(collectionUniqueList.Item(i), "UKI") > 0 Then
                    SplitWorksheet (collectionUniqueList.Item(i))
            End If
        Next i
 
        ActiveSheet.AutoFilterMode = False
 
    End With
 
Cleanup:
 
    Application.DisplayAlerts = True
    Set collectionUniqueList = Nothing
    Set wsSource = Nothing
    Set wsHelper = Nothing
 
End Sub
 
Private Sub Init_Unique_List_Collection(ByRef col As Collection, ByVal SourceWS_LastRow As Long)
 
    Dim LastRow As Long, RowNumber As Long
 
    wsSource.Range("B2:B" & SourceWS_LastRow).Copy wsHelper.Range("A1")
 
    With wsHelper
 
        If Len(Trim(.Range("A1").Value)) > 0 Then
 
            LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
 
            .Range("A1:A" & LastRow).RemoveDuplicates 1, xlNo
 
            LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
 
            .Range("A1:A" & LastRow).Sort .Range("A1"), Header:=xlNo
 
            LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
 
            On Error Resume Next
            For RowNumber = 1 To LastRow
                col.Add .Cells(RowNumber, "A").Value, CStr(.Cells(RowNumber, "A").Value)
            Next RowNumber
 
        End If
 
    End With
 
End Sub
 
Private Sub SplitWorksheet(ByVal Category_Name As Variant)
 
    Dim wbTarget As Workbook
 
    Set wbTarget = Workbooks.Add
 
    With wsSource
 
        With .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
            .AutoFilter .Range("B1").Column, Category_Name
 
            .Copy
 
            wbTarget.Worksheets(1).Paste
            wbTarget.Worksheets(1).Name = Category_Name
            Dim Today As String
            Today = Format(Now(), "yyyymmdd")
 
            wbTarget.SaveAs Target_Folder & "Inactive Employees" & Today & Category_Name & ".xlsx", 51
            wbTarget.Close False
 
        End With
 
    End With
 
    Set wbTarget = Nothing
 
End Sub
