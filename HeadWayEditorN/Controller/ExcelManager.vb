Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms.VisualStyles
Imports Excel = Microsoft.Office.Interop.Excel
Module ExcelManager

    Public Function CheckExcelAndUpdateWorkBook(filepath As String) As Integer
        Dim excelApp As New Excel.Application
        Dim excelWorkbook As Excel.Workbook = Nothing
        Try
            Dim FolderPath = Path.GetDirectoryName(filepath)
            Dim FileName = Path.GetFileNameWithoutExtension(filepath)
            Dim newFileName = IO.Path.GetFileNameWithoutExtension(filepath) + "-Updated-Type1.xlsx"
            excelWorkbook = excelApp.Workbooks.Open(filepath)
            Dim xws As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)


            Dim shapes = xws.Shapes
            'Dim cellA1 As String = CStr(xws.Range("A1").Value)
            'Dim cellC1 As String = CStr(xws.Range("C1").Value)
            'Dim cellA6 As String = CStr(xws.Range("A6").Value)
            'Dim cellA9 As String = CStr(xws.Range("A9").Value)

            'If CheckIfType1(cellA1, cellC1, cellA6, cellA9) Then
            '    excelWorkbook.SaveAs(Path.Combine(Path.Combine(FolderPath, "Original"), $"{FileName}.xlsx"))
            '    Return 1
            'Else
            '    Dim cellB1 As String = CStr(xws.Range("B1").Value)
            '    Dim cellD1 As String = CStr(xws.Range("D1").Value)
            '    Dim cellB6 As String = CStr(xws.Range("B6").Value)
            '    Dim cellB9 As String = CStr(xws.Range("B9").Value)
            '    If CheckIfType1(cellB1, cellD1, cellB6, cellB9) Then
            '        excelWorkbook.SaveAs(Path.Combine(Path.Combine(FolderPath, "Original"), $"{FileName}.xlsx"))
            '        Return 2
            '    End If
            'End If
            Dim type = GetTypeOfXls(xws)
            excelWorkbook.SaveAs(Path.Combine(Path.Combine(FolderPath, "Original"), $"{FileName}.xlsx"))
            Return type

        Catch ex As Exception
            ex.ToString()
            Return 0
        Finally
            If excelWorkbook IsNot Nothing Then
                excelWorkbook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            End If

            excelWorkbook = Nothing
            excelApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Private Function GetTypeOfXls(xws As Excel.Worksheet) As Int32
        Dim ret = 0
        ' Cells for Type1
        Dim cellA1 As String = CStr(xws.Range("A1").Value)
        Dim cellC1 As String = CStr(xws.Range("C1").Value)
        Dim cellA6 As String = CStr(xws.Range("A6").Value)
        Dim cellA9 As String = CStr(xws.Range("A9").Value)

        'Cells for Type2
        Dim cellB1 As String = CStr(xws.Range("B1").Value)
        Dim cellD1 As String = CStr(xws.Range("D1").Value)
        Dim cellB6 As String = CStr(xws.Range("B6").Value)
        Dim cellB9 As String = CStr(xws.Range("B9").Value)

        'Cells for Type3
        Dim cellJ1 As String = CStr(xws.Range("J1").Value)
        Dim cellM1 As String = CStr(xws.Range("M1").Value)

        Dim cellE1 As String = CStr(xws.Range("E1").Value)
        Dim cellG1 As String = CStr(xws.Range("G1").Value)
        Dim cellN1 As String = CStr(xws.Range("N1").Value)
        Dim cellQ1 As String = CStr(xws.Range("Q1").Value)

        ' Check Type1
        If cellA1 <> Nothing And cellC1 <> Nothing And cellA6 <> Nothing And cellA9 <> Nothing Then
            If cellA1.Trim() = "Property/Lease Info" And cellC1.Trim() = "Premises Component" And cellA6.Trim() = "Transaction Type" And cellA9.Trim() = "RR" Then
                Return 1
            End If
        End If

        'Check Type2
        If cellB1 <> Nothing And cellD1 <> Nothing And cellB6 <> Nothing And cellB9 <> Nothing Then
            If cellB1.Trim() = "Property/Lease Info" And cellD1.Trim() = "Premises Component" And cellB6.Trim() = "Transaction Type" And cellB9.Trim() = "RR" Then
                Return 2
            End If
        End If

        ' Check Type3
        If cellA1 <> Nothing And cellC1 <> Nothing And cellJ1 <> Nothing And cellM1 <> Nothing Then
            If cellA1.Trim() = "Address" And cellC1.Trim() = "Premises Component" And cellJ1.Trim() = "Rental pa" And cellM1.Trim() = "RR" Then
                Return 3
            End If
        End If

        If cellE1 <> Nothing And cellG1 <> Nothing And cellN1 <> Nothing And cellQ1 <> Nothing Then
            If cellE1.Trim() = "Address" And cellG1.Trim() = "Premises Component" And cellN1.Trim() = "Rental pa" And cellQ1.Trim() = "RR" Then
                Return 3
            End If
        End If
        Return ret
    End Function

    Public Function UpdateStyleExcel(filepath As String, type As Integer) As Integer

        Dim excelApp As New Excel.Application
        Dim excelWorkbook As Excel.Workbook = Nothing

        Try

            Dim newFileName = IO.Path.GetFileNameWithoutExtension(filepath) + "-Updated-Type1.xlsx"
            Dim folder = Path.Combine(Path.GetDirectoryName(filepath), "Edited")
            Dim logfolder = Path.Combine(Path.GetDirectoryName(filepath), "Logs")
            Dim newfile = Path.Combine(folder, newFileName)

            Dim sbuilder = ""
            sbuilder = sbuilder + $"originalfile =  {filepath}" + vbNewLine
            sbuilder = sbuilder + $"newfile = {newfile}" + vbNewLine
            sbuilder = sbuilder + vbNewLine

            Dim startdt = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            sbuilder = sbuilder + $"startime = {startdt}" + vbNewLine
            ' Open excel
            excelWorkbook = excelApp.Workbooks.Open(filepath)
            Dim xws As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

            If type = 1 Or type = 2 Then
                sbuilder = ProcessType1_2(xws, type, sbuilder)
            ElseIf type = 3 Then
                sbuilder = ProcessType3(xws, type, sbuilder)
            End If


            Dim enddt = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            sbuilder = sbuilder + $"endtime = {enddt}" + vbNewLine
            Dim logfilepath = $"{logfolder}/LOG-{Path.GetFileNameWithoutExtension(filepath)}.txt"
            WriteLogging(sbuilder, logfilepath)
            excelWorkbook.SaveAs(newfile)

            Return 0

        Catch ex As Exception
            ex.ToString()
            Return Nothing
        Finally
            If excelWorkbook IsNot Nothing Then
                excelWorkbook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            End If

            excelWorkbook = Nothing
            excelApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Public Function ProcessType1_2(xws As Excel.Worksheet, type As Integer, sbuilder As String) As String
        Dim resizeIndexes = New List(Of Integer)
        Dim shapes = xws.Shapes
        Dim ra1 = xws.Range("A1")
        Dim removeIndexes = New List(Of Integer)

        Dim paIndex = 9
        Dim rrIndex = 2

        If type = 1 Then
            paIndex = 8
            rrIndex = 1
        End If

        ' Photo Process
        If CStr(ra1.Value) = "Photo" Then
            Dim wa1 = ra1.Width
            If ra1.MergeCells Then
                wa1 = ra1.MergeArea.Width
            End If
            Dim www As Integer = 0

            For Each shape As Excel.Shape In shapes
                If shape.Width = wa1 Then
                    Continue For
                Else
                    If Not resizeIndexes.Contains(shape.TopLeftCell.Row) Then
                        If Not resizeIndexes.Contains(shape.TopLeftCell.Row - 1) Then
                            resizeIndexes.Add(shape.TopLeftCell.Row)
                        End If
                    End If
                End If
            Next
            xws.Range("A1").Activate()
        Else
            sbuilder = sbuilder + vbNewLine
        End If

        ' Remove Blanks
        Dim lastRow As Integer = xws.Range("D5000").End(Excel.XlDirection.xlUp).Row
        Dim startRow = 1
        While startRow < lastRow
            Dim cellB1 = DirectCast(xws.Cells(startRow, rrIndex), Excel.Range).Value2
            If cellB1 = "RR" Then
                Dim isvalue = DirectCast(xws.Cells(startRow - 3, rrIndex + 1), Excel.Range).Value2
                If isvalue <> "" Then
                    Dim dindex = startRow
                    Dim aa = DirectCast(xws.Cells(dindex + 1, paIndex), Excel.Range).Value2

                    While dindex < lastRow And aa <> "pa"
                        aa = DirectCast(xws.Cells(dindex + 1, paIndex), Excel.Range).Value2
                        If aa = "pa" Then
                            Continue While
                        End If
                        removeIndexes.Add(dindex + 1)
                        dindex += 1
                    End While
                    startRow = dindex
                End If
            End If

            If cellB1 = "Address" Then
                Dim cellc = DirectCast(xws.Cells(startRow, rrIndex + 1), Excel.Range).Value2
                If cellc = "" Then
                    Dim dindex = startRow
                    removeIndexes.Add(dindex - 1)
                    removeIndexes.Add(dindex)
                    Dim aa = DirectCast(xws.Cells(dindex + 1, paIndex), Excel.Range).Value2
                    While dindex < lastRow And aa <> "pa"
                        removeIndexes.Add(dindex + 1)
                        aa = DirectCast(xws.Cells(dindex + 1, paIndex), Excel.Range).Value2
                        dindex += 1
                    End While
                    removeIndexes.Add(dindex)
                    removeIndexes.Add(dindex + 1)
                    startRow = dindex
                End If
            End If
            startRow += 1
        End While
        removeIndexes.Reverse()


        For Each index As Integer In removeIndexes
            DirectCast(xws.Cells(index, 1), Excel.Range).EntireRow.Delete()
        Next
        sbuilder = sbuilder + $"rows deleted = [{removeIndexes.Count}]" + vbNewLine

        Dim drows = ""
        Dim i = 0
        For Each index As Integer In removeIndexes
            If i = removeIndexes.Count - 1 Then
                drows = drows + $"{index}"
            Else
                drows = drows + $"{index},"
            End If
            i = i + 1

        Next

        sbuilder = sbuilder + $"rows were = [{drows}]" + vbNewLine


        resizeIndexes.Sort()

        Dim resizeIndexes1 = New List(Of Integer)
        For Each ind As Integer In resizeIndexes
            If Not resizeIndexes1.Contains(ind - 1) And Not removeIndexes.Contains(ind) Then
                resizeIndexes1.Add(ind)
            End If
        Next

        Dim irows = ""
        i = 0
        For Each index1 As Integer In resizeIndexes1
            If i = resizeIndexes1.Count - 1 Then
                irows = irows + $"{index1}"
            Else
                irows = irows + $"{index1},"
            End If
            i = i + 1
        Next

        ' Fixed the Images
        If CStr(ra1.Value) = "Photo" Then
            Dim wa1 = ra1.Width
            If ra1.MergeCells Then
                wa1 = ra1.MergeArea.Width
            End If
            For Each shape As Excel.Shape In shapes
                If shape.Width = wa1 Then
                    Continue For
                Else
                    shape.Width = wa1
                End If
            Next
            xws.Range("A1").Activate()
        Else
            sbuilder = sbuilder + vbNewLine
        End If

        sbuilder = sbuilder + $"images resized = [{resizeIndexes1.Count}]" + vbNewLine
        sbuilder = sbuilder + $"images were = [{irows}]" + vbNewLine
        Return sbuilder
    End Function

    Public Function ProcessType3(xws As Excel.Worksheet, type As Integer, sbuilder As String) As String
        Dim resizeIndexes = New List(Of Integer)
        Dim shapes = xws.Shapes
        Dim ra1 = xws.Range("A1")
        Dim removeIndexes = New List(Of Integer)
        Dim WhiteColor = 16777215

        ' Photo Process
        If CStr(ra1.Value) = "Photo" Then
            Dim wa1 = ra1.Width
            If ra1.MergeCells Then
                wa1 = ra1.MergeArea.Width
            End If
            Dim www As Integer = 0

            For Each shape As Excel.Shape In shapes
                If shape.Width = wa1 Then
                    Continue For
                Else
                    If Not resizeIndexes.Contains(shape.TopLeftCell.Row) Then
                        If Not resizeIndexes.Contains(shape.TopLeftCell.Row - 1) Then
                            resizeIndexes.Add(shape.TopLeftCell.Row)
                        End If
                    End If
                End If
            Next
            xws.Range("A1").Activate()
        Else
            sbuilder = sbuilder + vbNewLine
        End If

        ' Remove Blanks
        Dim startRow = 2
        Dim height = ra1.Height
        Dim lastRow As Integer = xws.UsedRange.Rows.Count

        While startRow <= lastRow
            Dim color = DirectCast(xws.Cells(startRow, 1), Excel.Range).Interior.Color
            ' Check If White Color
            If color <> WhiteColor Then
                removeIndexes.Add(startRow)
            End If
            startRow = startRow + 1
        End While
        removeIndexes.Reverse()

        For Each index As Integer In removeIndexes
            DirectCast(xws.Cells(index, 1), Excel.Range).EntireRow.Delete()
        Next
        sbuilder = sbuilder + $"rows deleted = [{removeIndexes.Count}]" + vbNewLine

        Dim drows = ""
        Dim i = 0
        For Each index As Integer In removeIndexes
            If i = removeIndexes.Count - 1 Then
                drows = drows + $"{index}"
            Else
                drows = drows + $"{index},"
            End If
            i = i + 1

        Next

        sbuilder = sbuilder + $"rows were = [{drows}]" + vbNewLine


        resizeIndexes.Sort()

        Dim resizeIndexes1 = New List(Of Integer)

        For Each ind As Integer In resizeIndexes
            If Not resizeIndexes1.Contains(ind - 1) And Not removeIndexes.Contains(ind) Then
                resizeIndexes1.Add(ind)
            End If
        Next

        Dim irows = ""
        For Each index1 As Integer In resizeIndexes1
            If i = resizeIndexes1.Count - 1 Then
                irows = irows + $"{index1}"
            Else
                irows = irows + $"{index1},"
            End If
            i = i + 1
        Next

        ' Fixed the Images
        If CStr(ra1.Value) = "Photo" Then
            Dim wa1 = ra1.Width
            If ra1.MergeCells Then
                wa1 = ra1.MergeArea.Width
            End If
            For Each shape As Excel.Shape In shapes
                If shape.Width = wa1 Then
                    Continue For
                Else
                    shape.Width = wa1
                End If
            Next
            xws.Range("A1").Activate()
        Else
            sbuilder = sbuilder + vbNewLine
        End If

        sbuilder = sbuilder + $"images resized = [{resizeIndexes1.Count}]" + vbNewLine
        sbuilder = sbuilder + $"images were = [{irows}]" + vbNewLine
        Return sbuilder
    End Function
End Module
