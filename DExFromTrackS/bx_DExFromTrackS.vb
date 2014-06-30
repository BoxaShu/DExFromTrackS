
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface

Imports System.Text
Imports System.IO

Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.Serialization


'http://code.google.com/p/excellibrary/
'Imports ExcelLibrary.SpreadSheet

'http://closedxml.codeplex.com/
'Imports ClosedXML.Excel

'http://epplus.codeplex.com/documentation
'Imports OfficeOpenXml


Public Class acad__boxashu

    Dim _markers As DBObjectCollection = Nothing
    '    'http://adn-cis.org/ispolzovanie-tranzitnoj-grafiki.html
    '    'Использование транзитной графики

    <CommandMethod("bx_DExFromTrackS")> _
    Public Sub bx_DExFromTrackS()
        ' Получениеn текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        'Список описаний сечений
        Dim sec_list As New List(Of Section)
        'Список описаний сборок
        Dim CageList As New List(Of Cage)
        'Список трек линий
        Dim tl_list As New List(Of trackLine)
        'Итоговый список элементов для вывода.
        'Dim el_list As New Dictionary(Of String, Double)
        Dim el_list As New SortedList(Of String, Double)




        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            'Выбираем только атрибуты
            Dim acTypValAr(0) As TypedValue
            acTypValAr.SetValue(New TypedValue(DxfCode.Start, "MTEXT,LINE"), 0)
            '' Назначение критериев фильтра объекту SelectionFilter
            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
            Dim acSSPromptOpt As PromptSelectionOptions = New PromptSelectionOptions()
            acSSPromptOpt.AllowDuplicates = False
            acSSPromptOpt.MessageForAdding = ControlChars.CrLf & "Выберите описания сечений (MTEXT) и трек линии (LINE): "

            '' Запрос выбора объектов в области чертежа
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(acSSPromptOpt, acSelFtr)
            '' Если статус запроса равен OK, объекты выбраны
            If acSSPrompt.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim acSSet As SelectionSet = acSSPrompt.Value
            '' Перебор объектов в наборе
            For Each acSSObj As SelectedObject In acSSet
                '' Проверка, нужно убедится в правильности полученного объекта
                If Not IsDBNull(acSSObj) Then
                    '' Открытие объекта для записи
                    Dim acEnt As Entity = CType(acTrans.GetObject(acSSObj.ObjectId, _
                                                            OpenMode.ForRead), Entity)
                    If Not IsDBNull(acEnt) Then

                        If TypeOf acEnt Is Line Then
                            Dim acLine As Line = CType(acEnt, Line)
                            'Тестовая линия
                            Dim tl As trackLine = New trackLine
                            tl.Name = acLine.ColorIndex.ToString
                            tl.Length = acLine.Length
                            tl.ObjID = acLine.ObjectId

                            tl_list.Add(tl)
                        End If
                        If TypeOf acEnt Is MText Then
                            Dim acMtext As MText = CType(acEnt, MText)
                            'очистка текста от форматирования
                            Dim te As TextEditor = TextEditor.CreateTextEditor(acMtext)
                            ' Просто выбираем всё и удаляем форматирование
                            te.SelectAll()
                            te.Selection.RemoveAllFormatting()
                            ' Не забываем сохранить результаты
                            te.Close(TextEditor.ExitStatus.ExitSave)

                            Dim xlmString As String = "<?xml version=" & ControlChars.Quote & "1.0" & ControlChars.Quote & _
                            " encoding=" & ControlChars.Quote & "utf-16" & ControlChars.Quote & "?>"
                            xlmString = xlmString & ControlChars.CrLf & acMtext.Contents.Replace("\P", ControlChars.CrLf)


                            'Вот тут нужна проверка xml на соответствие схеме
                            'Нужно поискать все схемы в текущем каталоге и проверять на соответствие.
                            'http://msdn.microsoft.com/ru-ru/library/system.xml.schema.xmlschemaset%28v=vs.110%29.aspx



                            Dim productsXML As XElement = XDocument.Parse(xlmString).Root
                            'создаем reader
                            Dim reader As StringReader = New StringReader(xlmString)

                            If productsXML.Name = "Section" Then
                                'создаем XmlSerializer
                                Dim dsr As XmlSerializer = New XmlSerializer(GetType(Section))
                                'десериализуем 
                                Dim clone As Section = CType(dsr.Deserialize(reader), Section)

                                If sec_list.Contains(clone) = True Then
                                    acEd.WriteMessage(ControlChars.CrLf & "Присутствует двойное описание сечения " & clone.Name)
                                Else
                                    sec_list.Add(clone)
                                End If
                            End If

                            If productsXML.Name = "Cage" Then
                                'создаем XmlSerializer
                                Dim dsr As XmlSerializer = New XmlSerializer(GetType(Cage))
                                'десериализуем 
                                Dim clone As Cage = CType(dsr.Deserialize(reader), Cage)

                                If CageList.Contains(clone) = True Then
                                    acEd.WriteMessage(ControlChars.CrLf & "Присутствует двойное описание элемента " & clone.Name)
                                Else
                                    CageList.Add(clone)
                                End If
                            End If
                        End If
                    End If

                End If
            Next
            ' Сохранение нового объекта в базе данных
            acTrans.Commit()
            ' Очистка транзакции
        End Using


        'тут передираем все линии
        For Each i As trackLine In tl_list
            'тут обработать если описание не найдено
            'Dim aswerr As Integer = sec_list.FindIndex(Function(q) q.Name = i.Name)
            Dim CS_1 As Section
            If sec_list.FindIndex(Function(q) q.Name = i.Name) >= 0 Then
                CS_1 = sec_list.Item(sec_list.FindIndex(Function(q) q.Name = i.Name))
            Else
                acEd.WriteMessage(ControlChars.CrLf & "В выборке присутствует не описанная ТрекЛиния. Выполнение программы остановлено")
                'Пометить маркером не обработанную линию


                '''''''''''''''marker(i.ObjID)



                Exit Sub
            End If

            'тут блок подсчета элементов
            For Each q As Along In CS_1.Along_List
                If el_list.ContainsKey(q.Name) Then
                    Dim tempValve As Double = el_list.Item(q.Name)
                    el_list.Item(q.Name) = tempValve + (i.Length + q.Add) * q.Count
                Else
                    el_list.Add(q.Name, (i.Length + q.Add) * q.Count)
                End If
            Next
            For Each q As Across In CS_1.Across_List
                If el_list.ContainsKey(q.Name) Then
                    Dim tempValve As Double = el_list.Item(q.Name)
                    el_list.Item(q.Name) = tempValve + (i.Length / q.Steps) + q.Add
                Else
                    el_list.Add(q.Name, (i.Length / q.Steps) + q.Add)
                End If
            Next
        Next


        'Вывод результатов
        acEd.WriteMessage(ControlChars.CrLf & "Спецификация.")
        For Each i As KeyValuePair(Of String, Double) In el_list
            acEd.WriteMessage(ControlChars.CrLf & i.Key & " кол: " & i.Value)
        Next

        If CageList.Count > 0 Then

            acEd.WriteMessage(ControlChars.CrLf & "")
            acEd.WriteMessage(ControlChars.CrLf & "Сборочные единицы.")
            For Each i As Cage In CageList
                Dim tempString As String = ""
                If el_list.ContainsKey(i.Name) Then
                    tempString = " кол: " & el_list.Item(i.Name).ToString
                End If

                acEd.WriteMessage(ControlChars.CrLf & i.Name & tempString)
                For Each j As Along In i.Along_List
                    acEd.WriteMessage(ControlChars.CrLf & j.Name & " кол: " & j.Count)
                Next
            Next

        End If


        'создаем сериалайзер
        Dim list As New List(Of String)
        For Each i As KeyValuePair(Of String, Double) In el_list
            list.Add(i.Key & " " & i.Value.ToString)
        Next



        Dim sr As XmlSerializer = New XmlSerializer(list.GetType())
        Dim ns As New XmlSerializerNamespaces()
        ns.Add("", "")
        Dim settings As New XmlWriterSettings()
        settings.OmitXmlDeclaration = True
        'создаем writer, в который будет происходить сериализация
        Dim sb As StringBuilder = New StringBuilder()
        Dim w As StringWriter = New StringWriter(sb, System.Globalization.CultureInfo.InvariantCulture)
        'сериализуем
        sr.Serialize(w, list, ns)
        'получаем строку Xml
        Dim xml As String = sb.ToString()
        acEd.WriteMessage(ControlChars.CrLf)
        acEd.WriteMessage(xml)

        'тут Вывод результатов в excel
        'toXLS(el_list, CageList)

    End Sub

    '<CommandMethod("bx_toXLS")> _
    'Public Sub bx_toXLS()
    '    'Sub toXLS(ByRef el_list As SortedList(Of String, Double), ByRef CageList As List(Of Cage))

    '    'http://code.google.com/p/excellibrary/
    '    ' Получениеn текущего документа и базы данных
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim acEd As Editor = acDoc.Editor


    '    'Так же необходима процедура определения пути и имени файла
    '    'Имя файла
    '    Dim dwgName As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("DwgName")
    '    'Папка в которой находится файл
    '    Dim dwgPath As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("DWGPREFIX")

    '    Dim DllPath As String = My.Application.Info.DirectoryPath
    '    Dim outList As DirectoryInfo = New System.IO.DirectoryInfo(DllPath) 'System.IO.DirectoryInfo(DllPath)


    '    'Сначала определим будм ли выводить!
    '    'Dim getWhichEntityOptions As PromptKeywordOptions = New PromptKeywordOptions(ControlChars.Lf & "Вывести спецификацию в Excel? [Yes/No] : ", "Yes No")
    '    Dim getWhichEntityOptions As PromptKeywordOptions = New PromptKeywordOptions(ControlChars.Lf & "Вывести спецификацию в Excel?")
    '    'getWhichEntityOptions.Keywords.Add("Yes")
    '    'getWhichEntityOptions.Keywords.Add("No")

    '    For Each i As IO.FileInfo In outList.GetFiles("out*.xml")
    '        getWhichEntityOptions.Keywords.Add(i.Name)
    '    Next

    '    'Получить данные
    '    Dim getWhichEntityResult As PromptResult = acDoc.Editor.GetKeywords(getWhichEntityOptions)
    '    ' Если все ОК
    '    If (getWhichEntityResult.Status <> PromptStatus.OK) Then
    '        'Проверяем для того чтобы увидеть какой тип Объекта был введен
    '        Exit Sub
    '    End If

    '    'getWhichEntityResult.StringResult
    '    Dim inStr As String = "Бетон В25 W6 F150"

    '    Dim fileContents As String
    '    fileContents = My.Computer.FileSystem.ReadAllText(DllPath & "\" & getWhichEntityResult.StringResult)


    '    Dim productsXML As XElement = XDocument.Parse(fileContents).Root
    '    Dim Kr_1 As String = productsXML.FirstAttribute
    '    Dim Kr_2 As String = productsXML.LastAttribute
    '    Dim list As New Dictionary(Of String, String)
    '    For Each i As XElement In productsXML.Nodes
    '        ' i - это вложение like
    '        For Each q As XElement In i.Nodes
    '            'j-внутри like 
    '            list.Add(q.Name.ToString, q.Value)
    '        Next

    '        'For Each q As Xml.Linq.XAttribute In i.Attributes
    '        '    list.Add(q.Name.ToString, q.Value)
    '        'Next
    '    Next




    '    Dim FullFilePath As String = dwgPath & dwgName & ".xlsx"
    '    Dim newFile As FileInfo = New FileInfo(FullFilePath)
    '    Using ExcelPackage As ExcelPackage = New ExcelPackage(newFile)
    '        Dim excelWorksheet As ExcelWorksheet = ExcelPackage.Workbook.Worksheets.Add("Спецификация")

    '        excelWorksheet.Cells(1, 1).Value = 1
    '        excelWorksheet.Cells(1, 2).Value = 1
    '        excelWorksheet.Cells(1, 3).FormulaR1C1 = "=R1C1+R1C2"
    '        ExcelPackage.Save()
    '    End Using




    '    ''Вывод результатов
    '    'acEd.WriteMessage(ControlChars.CrLf & "Спецификация.")
    '    'For Each i As KeyValuePair(Of String, Double) In el_list
    '    '    acEd.WriteMessage(ControlChars.CrLf & i.Key & " кол: " & i.Value)
    '    'Next

    '    'If CageList.Count > 0 Then

    '    '    acEd.WriteMessage(ControlChars.CrLf & "")
    '    '    acEd.WriteMessage(ControlChars.CrLf & "Сборочные единицы.")
    '    '    For Each i As Cage In CageList
    '    '        Dim tempString As String = ""
    '    '        If el_list.ContainsKey(i.Name) Then
    '    '            tempString = " кол: " & el_list.Item(i.Name).ToString
    '    '        End If

    '    '        acEd.WriteMessage(ControlChars.CrLf & i.Name & tempString)
    '    '        For Each j As Along In i.Along_List
    '    '            acEd.WriteMessage(ControlChars.CrLf & j.Name & " кол: " & j.Count)
    '    '        Next
    '    '    Next

    '    'End If


    '    Process.Start(FullFilePath)
    'End Sub




    'Sub marker(ByRef ObjID As ObjectId)
    '    Dim activeDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = activeDoc.Database
    '    Dim ed As Editor = activeDoc.Editor

    '    ClearTransientGraphics()
    '    _markers = New DBObjectCollection
    '    Using tr As Transaction = db.TransactionManager.StartTransaction()
    '        Dim line As Line = TryCast(tr.GetObject(ObjID, OpenMode.ForRead), Line)
    '        Dim cnt As Integer = 0

    '        Dim pt As Point3d = function_lib.CenterPoint(line.StartPoint, line.EndPoint)
    '        Dim marker As New Circle(pt, Vector3d.ZAxis, 50)
    '        marker.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 255, 0)

    '        _markers.Add(marker)
    '        Dim intCol As New IntegerCollection()
    '        Dim tm As TransientManager = TransientManager.CurrentTransientManager
    '        tm.AddTransient(marker, TransientDrawingMode.Highlight, 128, intCol)
    '        ed.WriteMessage(vbLf + pt.ToString())

    '        System.Math.Max(System.Threading.Interlocked.Increment(cnt), cnt - 1)

    '        tr.Commit()
    '    End Using
    'End Sub


    Sub ClearTransientGraphics()
        Dim tm As TransientManager = TransientManager.CurrentTransientManager
        Dim intCol As IntegerCollection = New IntegerCollection()
        If _markers <> Nothing Then
            For Each marker As DBObject In _markers
                tm.EraseTransient(marker, intCol)
                marker.Dispose()
            Next
        End If
    End Sub

End Class


'Sub toXLS(ByRef el_list As SortedList(Of String, Double))

'    'http://www.carlosag.net/tools/excelxmlwriter/
'    ' Получениеn текущего документа и базы данных
'    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
'    Dim acCurDb As Database = acDoc.Database
'    Dim acEd As Editor = acDoc.Editor



'    Dim book As Workbook = New Workbook()
'    Dim sheet As Worksheet = book.Worksheets.Add("Спецификация")
'    Dim row As WorksheetRow = sheet.Table.Rows.Add()
'    Dim cell As WorksheetCell ' = row.Cells.Add("Header 3")

'    row.Cells.Add("Спецификация")
'    cell = row.Cells.Add()
'    cell.MergeAcross = 11

'    row = sheet.Table.Rows.Add()


'    For Each i As KeyValuePair(Of String, Double) In el_list
'        Dim strArray() As String = i.Key.Split("|")

'        'acEd.WriteMessage(ControlChars.CrLf & i.Key & " кол: " & i.Value)
'    Next


'    Dim filename As String = "D:\test.xls"
'    book.Save(filename)
'    Process.Start(filename)
'End Sub