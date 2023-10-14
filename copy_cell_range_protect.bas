Attribute VB_Name = "Module1"
'Sub Test()
'    Dim SrcSh As Worksheet, DstSh As Worksheet
'    Set SrcSh = ThisWorkbook.Worksheets("temp")
'    Set DstSh = ThisWorkbook.Worksheets("Sheet4")
'    Call CopyCellRange(SrcSh, DstSh)
'End Sub
'#############################################################################
' �w��Z���̈��ʃV�[�g�̈�ɃR�s�[(�ی�Ȃ��j
'
'4�s1�񂩂�ŏI�sV��܂ł̗̈���R�s�[
'�@copy_cell_range
'#############################################################################
'Sub CopyCellRange()
'    Dim SrcSh As Worksheet, DstSh As Worksheet
'    Set SrcSh = ThisWorkbook.Worksheets("temp")
'    Set DstSh = ThisWorkbook.Worksheets("Sheet2")
Sub CopyCellRange(ByRef SrcSh As Worksheet, DstSh As Worksheet)
    Dim StrtRow As Long, StrtCol As Long, LastRow As Long, ColV As Long
    Dim TargetCol As String
    Dim SrcRange As Range, DstRange As Range
    StrtRow = 4
    StrtCol = 1
    TargetCol = "T"
    ColV = SrcSh.Range(TargetCol & "1").Column
    LastRow = SrcSh.Cells(SrcSh.Rows.Count, 1).End(xlUp).Row
'    MsgBox "�� " & TargetCol & " �̗�ԍ�: " & ColV
'    MsgBox "�ŏI�s�̍s�ԍ�: " & LastRow
    Set SrcRange = SrcSh.Range(SrcSh.Cells(StrtRow, StrtCol), SrcSh.Cells(LastRow, ColV))
    Set DstRange = DstSh.Range(DstSh.Cells(StrtRow, StrtCol), DstSh.Cells(LastRow, ColV))
    SrcRange.Copy DstRange
    Application.CutCopyMode = False
End Sub
Sub Test()
    Dim SrcSh As Worksheet, DstSh As Worksheet
    Set SrcSh = ThisWorkbook.Worksheets("temp")
    Set DstSh = ThisWorkbook.Worksheets("Sheet1")
    Call CopyCellRangeProtect(SrcSh, DstSh)
End Sub
'#############################################################################
' �w��Z���̈��ʃV�[�g�̈�ɃR�s�[(�ی삠��j
'
'4�s1�񂩂�ŏI�sV��܂ł̗̈���R�s�[
'�@copy_cell_range_protect
'#############################################################################
Sub CopyCellRangeProtect(ByRef SrcSh As Worksheet, DstSh As Worksheet)
    Dim StrtRow As Long, StrtCol As Long, LastRow As Long, ColV As Long
    Dim TargetCol As String
    Dim SrcRange As Range, DstRange As Range
    StrtRow = 4
    StrtCol = 1
    TargetCol = "T"
    ColV = SrcSh.Range(TargetCol & "1").Column
    LastRow = SrcSh.Cells(SrcSh.Rows.Count, 1).End(xlUp).Row
'    MsgBox "�� " & TargetCol & " �̗�ԍ�: " & ColV
'    MsgBox "�ŏI�s�̍s�ԍ�: " & LastRow
    Set SrcRange = SrcSh.Range(SrcSh.Cells(StrtRow, StrtCol), SrcSh.Cells(LastRow, ColV))
    Set DstRange = DstSh.Range(DstSh.Cells(StrtRow, StrtCol), DstSh.Cells(LastRow, ColV))
    SrcRange.Copy DstRange
    Application.CutCopyMode = False
    DstSh.Unprotect
    DstSh.Cells.Locked = False
    DstRange.Locked = True
    DstSh.Protect
End Sub
