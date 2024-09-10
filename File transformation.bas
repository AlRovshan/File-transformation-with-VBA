Attribute VB_Name = "Module1"
Sub impt()
Dim destinationCell, targetRange, sourceRange, cell, rowsToDelete As Range
Dim additionalSymbol, replacementSymbol, Name, savePath, group As String
Dim symbolToReplace, Number As Variant
Dim previousRow, i, j, ind As Integer
Dim row_count, newRowN As Long
'Получение названия списка
    group = Left(InputBox("Введите название списка"), 100)

'Проверка последовательности записей и удаление лишних строк
For row_count = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
    If Not IsNumeric(ActiveSheet.UsedRange.Cells(row_count, 1).Value) Or _
    IsEmpty(ActiveSheet.UsedRange.Cells(row_count, 1).Value) Then
        ActiveSheet.UsedRange.Cells(row_count, 1).EntireRow.Delete
    End If
Next row_count

For row_count = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
    If IsEmpty(ActiveSheet.UsedRange.Cells(row_count, 2).Value) Then
        ActiveSheet.UsedRange.Cells(row_count, 2).EntireRow.Delete
    End If
Next row_count

newRowN = 1
For row_count = 1 To ActiveSheet.UsedRange.Rows.Count
    If Not IsEmpty(ActiveSheet.UsedRange.Cells(row_count, 1).Value) Then
        ActiveSheet.UsedRange.Cells(row_count, 1).Value = newRowN
        newRowN = newRowN + 1
    End If
Next row_count

'Удаление лишних пробелов
With ActiveSheet.UsedRange
    .Value = Application.Trim(.Value)
End With

'Удаление лишних символов
    
    symbolsToReplace = Array("@", "#", "*", "-", "(", ")", " ", "_")
    replacementSymbol = ""
    For Each cell In ActiveSheet.UsedRange.Cells
        If cell.Column >= 3 And cell.Column <= 8 Then
            For Each sym In symbolsToReplace
                If InStr(1, cell.Value, sym) > 0 Then
                    cell.Value = Replace(cell.Value, sym, replacementSymbol)
                End If
        Next sym
        If Not IsNumeric(cell.Value) Then
            cell.ClearContents
        End If
        End If
'Замена цифры 8 на +7
        If cell.Column = 8 Then
            If Not IsEmpty(cell.Value) And Left(cell.Value, 1) = "8" Then
                cell.NumberFormat = "@"
                cell.Value = "+7" & Mid(cell.Value, 2)
            End If
        End If
    Next cell
    
'Создание шабона
Sheets.Add
additionalSymbol = ";"
i = 1
j = 0
previousRow = 1
Number = Sheets("Лист1").Cells(1, 1).Value
Name = Sheets("Лист1").Cells(1, 2).Value
Set destinationCell = ActiveCell.Offset(1, 0)

    For Each cell In Sheets("Лист1").UsedRange.Cells
        If cell.Column >= 3 And cell.Column <= 9 And Not IsEmpty(cell.Value) Then
            destinationCell.Value = Number & additionalSymbol & Name & additionalSymbol & _
            0 & additionalSymbol & "phone" & additionalSymbol
            'Назначение параметров Record number и Full name
                If cell.Row <> previousRow Then
                    i = i + 1
                    previousRow = cell.Row
                End If
                destinationCell.Value = Replace(destinationCell.Value, Number, Sheets("Лист1").Cells(i, 1).Value)
                destinationCell.Value = Replace(destinationCell.Value, Name, Left(Sheets("Лист1").Cells(i, 2).Value, 256))
                'Назначение параметра Contact oreder
                If cell.Column >= 4 And cell.Column <= 7 Then
                    j = j + 1
                Else
                    j = 0
                End If
                'Преобразавание ячеек с рабочим номером
                If cell.Column >= 3 And cell.Column <= 5 Then
                    destinationCell.Value = destinationCell.Value & j & additionalSymbol & cell.Value & _
                    String(2, additionalSymbol) & group & additionalSymbol & 1 & String(3, additionalSymbol)
                Else
                    'Преобразования ячеек с мобильным номером
                    If cell.Column = 7 Then
                        destinationCell.Value = destinationCell.Value & j & additionalSymbol & cell.Value & _
                        String(2, additionalSymbol) & group & additionalSymbol & 2 & String(3, additionalSymbol)
                    Else
                    'Преобразования ячеек с домашним номером
                        If cell.Column = 6 Then
                            destinationCell.Value = destinationCell.Value & j & additionalSymbol & cell.Value & _
                            String(2, additionalSymbol) & group & additionalSymbol & 3 & String(3, additionalSymbol)
                        Else
                            'Преобразования ячеек с sms
                            destinationCell.Value = destinationCell.Value & 0 & additionalSymbol & cell.Value & _
                            String(2, additionalSymbol) & group & String(4, additionalSymbol)
                            If cell.Column = 8 Then
                                destinationCell.Value = Replace(destinationCell.Value, "phone", "sms")
                            Else
                                'Преобразования ячеек с email
                                If cell.Column = 9 Then
                                destinationCell.Value = Replace(destinationCell.Value, "phone", "email")
                                End If
                            End If
                        End If
                    End If
                End If
                'Переход на следующую строку
                Set destinationCell = destinationCell.Offset(1, 0)
            End If
    Next cell
'Добавление названий столбцов в первую строку
Range("A1").Value = "Record number;Full name;Priority;Contact type;Contact oreder;Contact value;Confirmation PIN;Groups;Contact subtype;Opt info;Ext id;Is owned by group"
'Сохранение данных в новом файле в формате csv utf-8
Application.DisplayAlerts = False
savePath = ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".csv"
    ActiveSheet.Copy
    ActiveSheet.SaveAs filename:=savePath, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWindow.Close
    
    ChangeFileCharset savePath, "utf-8", "Windows-1251"
    
    'Удаление всех лишних данных в книге
    ActiveSheet.Delete
    ActiveSheet.Cells.Delete Shift:=xlUp
    ThisWorkbook.Saved = True
    ActiveWindow.Close
    Application.DisplayAlerts = True

End Sub

Function ChangeFileCharset(ByVal filename$, ByVal DestCharset$, _
                           Optional ByVal SourceCharset$) As Boolean
    On Error Resume Next: Err.Clear
    With CreateObject("ADODB.Stream")
        .Type = 2
        If Len(SourceCharset$) Then .Charset = SourceCharset$
        .Open
        .LoadFromFile filename$
        fileContent$ = .ReadText
        .Close
        .Charset = DestCharset$
        .Open
        .WriteText fileContent$
        
        Dim binaryStream As Object
        Set binaryStream = CreateObject("ADODB.Stream")
        binaryStream.Type = 1
        binaryStream.Mode = 3
        binaryStream.Open
        
        .Position = 3
        .CopyTo binaryStream
        .Flush
        .Close
        binaryStream.SaveToFile filename$, 2
        binaryStream.Close
    End With
    ChangeFileCharset = Err = 0
End Function


