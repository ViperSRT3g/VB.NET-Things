Imports Microsoft.Office.Interop.Excel

Module Mod_ExcelAPIs
    Public Function GetLastRow(TSheet As Worksheet, ColumnNo As Long) As Long
        Return TSheet.Cells(TSheet.Rows.CountLarge, ColumnNo).End(-4162).Row
    End Function

    Public Function GetLastCol(TSheet As Worksheet, RowNo As Long) As Long
        Return TSheet.Cells(RowNo, TSheet.Columns.Count).End(-4159).Column
    End Function

    Public Function GetHeader(TSheet As Worksheet, HeaderRow As Long, HeaderStr As String) As Long
        Dim Header As Range = TSheet.Rows(HeaderRow).Find(HeaderStr, LookAt:=1)
        If Not Header Is Nothing Then Return Header.Column
    End Function

    Public Function GetHeaders(TSheet As Worksheet, HeaderRow As Long, Optional CaseSensitive As Boolean = False) As Dictionary(Of String, Long)
        Dim Output As New Dictionary(Of String, Long)
        For ColCounter As Long = 1 To GetLastCol(TSheet, HeaderRow)
            If CaseSensitive Then 'Headers are untouched
                Output(CStr(TSheet.Cells(HeaderRow, ColCounter).Value)) = ColCounter
            Else 'Headers are all Uppercase
                Output(UCase(CStr(TSheet.Cells(HeaderRow, ColCounter).Value))) = ColCounter
            End If
        Next ColCounter
        Return Output
    End Function

    Public Function Expand(ByRef Target As Range, ByVal Direction As XlDirection) As Range
        If Not Target Is Nothing Then Return Target.Parent.Range(Target, Target.End(Direction))
    End Function

    Public Function GetSheet(SheetName As String, TWB As Workbook) As Worksheet
        If SheetName.Length = 0 Then Exit Function
        If TWB Is Nothing Then Exit Function

        Dim Output As Worksheet
        Try
            Output = TWB.Worksheets(SheetName)
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        Finally
            If Output Is Nothing Then
                Output = TWB.Worksheets.Add(After:=TWB.Worksheets(TWB.Worksheets.Count))
                Output.Name = SheetName
            End If
        End Try
        Return Output
    End Function
End Module