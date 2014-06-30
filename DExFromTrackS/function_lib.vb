
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors


Module function_lib

#Region "вывод примитивов"
    Public Sub AddPoint(ByVal p1 As Point3d, ByVal colorIndex As Integer, ByVal type As Integer, ByVal size As Integer)

        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)
            '' Создание точки с координатой (4, 3, 0) в пространстве Модели
            Dim acPoint As DBPoint = New DBPoint(p1)

            acPoint.ColorIndex = CInt(colorIndex)
            acPoint.SetDatabaseDefaults()

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acPoint)
            acTrans.AddNewlyCreatedDBObject(acPoint, True)

            '' Установка стиля для всех объектов точек в чертеже
            acCurDb.Pdmode = type
            acCurDb.Pdsize = size

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub


    Sub AddPLine(ByVal GL_POLYs As List(Of Point3d), ByVal p3d As Point3d, ByVal closed As Boolean)

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков пространства Модели для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)

            '' Создание полилинии с двумя сегментами (3 точки)
            Dim acPoly As Polyline = New Polyline()
            acPoly.SetDatabaseDefaults()

            Dim q As Integer = 0
            For Each i As Point3d In GL_POLYs
                acPoly.AddVertexAt(q, New Point2d(i.X + p3d.X, i.Y + p3d.Y), 0, 0, 0) '37.000  21.950   0.000
                q = q + 1
            Next

            acPoly.Closed = closed

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub


    Sub AddLine(ByVal p1 As Point3d, ByVal p2 As Point3d, ByVal colorIndex As Integer)

        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков пространства Модели для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)

            '' Создание отрезка начинающегося в 5,5 и заканчивающегося в 12,3
            Dim acLine As Line = New Line(p1, p2)
            acLine.ColorIndex = colorIndex
            acLine.SetDatabaseDefaults()

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub

    Sub AddText(ByVal p1 As Point3d, ByVal str As String, ByVal Height As Double, ByVal colorIndex As Double)

        '' Устанавливаем текущий документ и базу данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Начинаем транзакцию
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открываем таблицу Блока для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
                                    OpenMode.ForRead)

            '' Открываем запись таблицы Блока пространство Модели (Model space) для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)

            '' Создаем однострочный текстовый объект
            Dim acText As DBText = New DBText()
            acText.SetDatabaseDefaults()
            acText.Position = p1 'New Point3d(2, 2, 0)
            acText.Height = Height
            acText.TextString = str
            acText.ColorIndex = CInt(colorIndex)
            'acText.AlignmentPoint = p1
            acText.HorizontalMode = TextHorizontalMode.TextCenter
            acText.VerticalMode = TextVerticalMode.TextVerticalMid
            acText.AlignmentPoint = p1
            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            '' Сохраняем изменения и закрываем транзакцию
            acTrans.Commit()
        End Using
    End Sub

    Public Sub crLayers(ByRef str As String)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        'Функция создания слоев
        'Начало транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            'Открываю таблицу слоев для чтения
            Dim acLyrTbl As LayerTable
            'тут возникает ошибка при отладке.
            acLyrTbl = CType(acTrans.GetObject(acCurDb.LayerTableId, _
                                         OpenMode.ForRead), LayerTable)
            'В этой переменной наименование слоя
            Dim sLayerName As String = str.ToString
            'Если этого слоя нет, то создаем его
            If acLyrTbl.Has(sLayerName) = False Then
                Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                'Создаем новый слой с заданными параметрами
                acLyrTblRec.Name = sLayerName
                acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 148)
                acLyrTblRec.Description = sLayerName + " создан программой"
                acLyrTblRec.LineWeight = LineWeight.LineWeight015

                acLyrTblRec.IsPlottable = True
                acLyrTblRec.IsOff = False
                acLyrTblRec.IsFrozen = False
                acLyrTblRec.IsLocked = False

                'Обновляем таблицу слоев для записи
                acLyrTbl.UpgradeOpen()

                'Добавляем новый слой в таблицу слоев
                acLyrTbl.Add(acLyrTblRec)
                acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
            End If

            'Сохранение изменений и завершение транзакции
            acTrans.Commit()
        End Using
    End Sub
#End Region


#Region "Геометрические функции"
    Public Function GetLength(ByVal varStart As Point3d, ByVal varEnd As Point3d) As Double
        Dim dblLen As Double
        On Error GoTo Err_Control
        dblLen = Math.Sqrt((varStart.X - varEnd.X) ^ 2 + _
        (varStart.Y - varEnd.Y) ^ 2 + _
        (varStart.Z - varEnd.Z) ^ 2)
        GetLength = dblLen
Exit_here:
        Exit Function
Err_Control:
        'MsgBox Err.Description
    End Function


    Function CenterPoint(ByVal pnt1 As Point3d, ByVal pnt2 As Point3d) As Point3d
        'Функция нахождения точки находящейся на равном удаление от двух заданных	    
        On Error GoTo Err_Control
        Dim A1 As Double = (pnt1.Y - pnt2.Y)
        Dim B1 As Double = (pnt2.X - pnt1.X)
        Dim C1 As Double = (pnt1.X * pnt2.Y - pnt2.X * pnt1.Y)
        Dim x_cen As Double = (pnt1.X + pnt2.X) / 2
        Dim y_cen As Double
        If Math.Round(B1, 3) <> 0 Then
            '(y1-y2)*x+(x2-x1)*y+(x1*y2-x2*y1)=0
            y_cen = (A1 * x_cen + C1) / -B1
        Else
            y_cen = (pnt1(1) + pnt2(1)) / 2
        End If

        'Возвращаю результат вычесления
        CenterPoint = New Point3d(x_cen, y_cen, 0)
Exit_here:
        Exit Function
Err_Control:
        MsgBox(Err.Description)
    End Function
#End Region


End Module
