Option Strict Off

Imports System.Drawing
Imports System.IO        ' Ver1.2.0で追加

Public Enum CellLineStyle
    Continuous = 1
    Dash = -4115
    DashDot = 4
    DashDotDot = 5
    Dot = -4118
    Ldouble = -4119
    None = -4142
    SlantDashDot = 13
End Enum

Public Enum CellBorderIndex
    DiagonalDown = 5
    DiagonalUp = 6
    EdgeLeft = 7
    EdgeTop = 8
    EdgeBottom = 9
    EdgeRight = 10
    InsideVertical = 11
    InsideHorizontal = 12
End Enum

Public Enum CellBorderWeight
    HairLine = 1
    Thin = 2
    Medium = -4138
    Thick = 4
End Enum

Public Class CellBorder

    Private _index As CellBorderIndex
    Private _color As Color
    Private _style As CellLineStyle
    Private _weight As CellBorderWeight

    Public Sub New(Optional ByVal vstyle As CellLineStyle = CellLineStyle.Continuous, _
                   Optional ByVal vweight As CellBorderWeight = CellBorderWeight.Thin, _
                   Optional ByVal vindex As CellBorderIndex = 0)

        _index = vindex
        _color = Color.Black
        _style = vstyle
        _weight = vweight
    End Sub

    Public Property Index() As CellBorderIndex
        Get
            Return _index
        End Get
        Set(ByVal value As CellBorderIndex)
            _index = value
        End Set
    End Property

    Public Property Color() As Color
        Get
            Return _color
        End Get
        Set(ByVal value As Color)
            _color = value
        End Set
    End Property

    Public Property LineStyle() As CellLineStyle
        Get
            Return _style
        End Get
        Set(ByVal value As CellLineStyle)
            _style = value
        End Set
    End Property

    Public Property Weight() As CellBorderWeight
        Get
            Return _weight
        End Get
        Set(ByVal value As CellBorderWeight)
            _weight = value
        End Set
    End Property

End Class

Public Class ExcelLB

    Private _excel As Object = Nothing

    ''' <summary>
    ''' コンストラクター
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        _excel = CreateObject("Excel.Application")
    End Sub

    ''' <summary> [ReleaseObject]：COMオブジェクトを解放するメソッド </summary>
    ''' <param name="target">COMオブジェクト</param>
    ''' <remarks></remarks>
    Public Sub ReleaseObject(ByVal target As Object)
        Try
            If Not target Is Nothing Then
                Do While System.Runtime.InteropServices.Marshal.ReleaseComObject(target) > 0
                    '
                Loop
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(target)
            End If
        Finally
            target = Nothing
        End Try
    End Sub

#Region "Excel本体"

    ''' <summary> [DisplayAlerts]プロパティ </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property DisplayAlerts() As Boolean
        Set(ByVal value As Boolean)
            _excel.DisplayAlerts = value
        End Set
    End Property

    ''' <summary> [ScreenUpdating]プロパティ </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property ScreenUpdating() As Boolean
        Set(ByVal value As Boolean)
            _excel.ScreenUpdating = value
        End Set
    End Property

    ''' <summary> [Visible]プロパティ </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property Visible() As Boolean
        Set(ByVal value As Boolean)
            _excel.Visible = value
        End Set
    End Property

    ''' <summary> [Dispose]：Excelオブジェクトを破棄するメソッド </summary>
    ''' <remarks></remarks>
    Public Sub Dispose()
        If Not _excel Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_excel)
            _excel = Nothing
        End If
    End Sub

    ''' <summary> [Quit]：Excelを終了するメソッド </summary>
    ''' <remarks></remarks>
    Public Sub Quit()
        Try
            _excel.Quit()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetStatusBar]：StatusBarを設定するメソッド </summary>
    ''' <param name="msg"></param>
    ''' <remarks></remarks>
    Public Sub SetStatusBar(ByVal msg As String)
        Try
            _excel.StatusBar = msg
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region "Workbook関係"

    ''' <summary> [Workbooks]プロパティ </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Workbooks() As Object
        Get
            Return _excel.Workbooks
        End Get
    End Property

    ''' <summary> [AddBook]：Workbookを追加するメソッド </summary>
    ''' <param name="books">Workbooks</param>
    ''' <returns>追加したBookを返す</returns>
    ''' <remarks></remarks>
    Public Function AddBook(ByVal books As Object) As Object
        Try
            Return books.Add()
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [OpenBook]：指定されたファイル名のWorkbookを開いて取得するメソッド </summary>
    ''' <param name="books">Workbooks</param>
    ''' <param name="filepath">ファイルパス</param>
    ''' <returns>開いたBookを返す</returns>
    ''' <remarks></remarks>
    Public Function OpenBook(ByVal books As Object, ByVal filepath As String) As Object
        Try
            Return books.Open(filepath)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [CloseBook]：WorkbookをCloseするメソッド </summary>
    ''' <param name="book">Workbook</param>
    ''' <remarks></remarks>
    Public Sub CloseBook(ByVal book As Object)
        Try
            book.Close()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [Save]：Workbookを上書き保存するメソッド </summary>
    ''' <param name="book">Workbook</param>
    ''' <remarks></remarks>
    Public Sub Save(ByVal book As Object)
        Try
            book.Save()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SaveAs]：Workbookに名前を付けて保存するメソッド </summary>
    ''' <param name="book">Workbook</param>
    ''' <param name="filepath">ファイルパス</param>
    ''' <remarks></remarks>
    Public Sub SaveAs(ByVal book As Object, ByVal filepath As String)

        Const NORMAL As Integer = -4143        ' Excel.XlFileFormat.xlWorkbookNormal
        Const XLSX As Integer = 51
        Const XLSM As Integer = 52
        Const XL8 As Integer = 56

        Try
            ' Ver1.2.0で以下の部分を変更
            If CType(_excel.version.ToString, Decimal) < 12 Then
                book.SaveAs(filepath, NORMAL)
            Else
                Dim ext = Path.GetExtension(filepath).ToLower()
                Dim format = XL8
                If ext = ".xlsx" Then
                    format = XLSX
                ElseIf ext = ".xlsm" Then
                    format = XLSM
                End If
                book.SaveAs(filepath, format)
            End If

        Catch ex As Exception
            Try
                book.SaveAs(filepath)
            Catch ex2 As Exception
                Throw
            End Try
        End Try
    End Sub

#End Region

#Region "Worksheet関係"

    ''' <summary> [CountSheet]：Worksheetの枚数を返すメソッド </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <returns>シート枚数</returns>
    ''' <remarks></remarks>
    Public Function CountSheet(ByVal sheets As Object) As Integer
        Try
            Return sheets.Count
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [ActiveSheet]：ActiveSheetを返すメソッド </summary>
    ''' <returns>ActiveSheet</returns>
    ''' <remarks></remarks>
    Public Function ActiveSheet() As Object
        Try
            Return _excel.ActiveSheet
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [AddSheet]：Worksheetを挿入するメソッド１ </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <param name="sheetname">挿入するWorksheetの名前</param>
    ''' <returns>Worksheet</returns>
    ''' <remarks>この方法では追加シートが最後尾に挿入される</remarks>
    Public Function AddSheet(ByVal sheets As Object, ByVal sheetname As String) As Object

        If Not ExistSheet(sheets, sheetname) Then
            Dim count As Integer = sheets.Count
            Dim sheet0 = sheets.Item(count)
            Dim sheet1 = sheets.Add(After:=sheet0)
            sheet1.Name = sheetname
            ReleaseObject(sheet0)
            ReleaseObject(sheet1)
            Return sheet1
        Else
            Dim msg = "[" & sheetname & "]シートは存在するので挿入できません。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [AddSheet]：Worksheetを挿入するメソッド２ </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <param name="sheetname">シート名</param>
    ''' <param name="index">挿入するシートの位置</param>
    ''' <returns>Worksheet</returns>
    ''' <remarks></remarks>
    Public Function AddSheet(ByVal sheets As Object, ByVal sheetname As String,
                             ByVal index As Integer) As Object

        Dim count As Integer = sheets.Count
        If count < index Then
            Dim msg = "引数[index] (= " & index.ToString & ") が不適切です。"
            Throw New ArgumentException(msg)
        End If

        If Not ExistSheet(sheets, sheetname) Then
            Dim sheet0 = sheets.Item(index)
            Dim sheet1 = sheets.Add(After:=sheet0)
            sheet1.Name = sheetname
            ReleaseObject(sheet0)
            ReleaseObject(sheet1)
            Return sheet1
        Else
            Dim msg = "[" & sheetname & "]シートは存在するので挿入できません。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [GetSheet]：Worksheetを取得するメソッド１ </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <param name="index">Sheet番号</param>
    ''' <returns>指定されたSheet番号のWorksheetを返す</returns>
    ''' <remarks></remarks>
    Public Function GetSheet(ByVal sheets As Object, ByVal index As Integer) As Object

        Dim count As Integer = sheets.Count
        If count < index Then
            Throw New ArgumentException("indexがシートの数を超えています。")
        Else
            Return sheets.Item(index)
        End If
    End Function

    ''' <summary> [GetSheet]：Worksheetを取得するメソッド２ </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <param name="sheetname">Sheet名</param>
    ''' <returns>指定されたSheet名のWorksheetを返す</returns>
    ''' <remarks></remarks>
    Public Function GetSheet(ByVal sheets As Object, ByVal sheetname As String) As Object

        If ExistSheet(sheets, sheetname) Then
            Return sheets.Item(sheetname)
        Else
            Dim msg = "存在しないシートが指定されました。(" & sheetname & ")"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [GetSheets]：WorkSheetsを取得するメソッド </summary>
    ''' <param name="book">Workbook</param>
    ''' <returns>Worksheetsを返す</returns>
    ''' <remarks></remarks>
    Public Function GetSheets(ByVal book As Object) As Object
        Try
            Return book.Worksheets
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [GetSheetName]：Worksheetの名前を取得するメソッド </summary>
    ''' <param name="sheet">Worksheet</param>
    ''' <returns>シート名</returns>
    ''' <remarks></remarks>
    Public Function GetSheetName(ByVal sheet As Object) As String
        Try
            Return sheet.Name
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [CopySheet]：Worksheetをコピーするメソッド </summary>
    ''' <param name="source">コピー元のWorksheet</param>
    ''' <param name="dest">コピー先位置決め用のWorksheet</param>
    ''' <param name="afterflag">True:destの後ろ、False:destの前にコピー</param>
    ''' <remarks></remarks>
    Public Sub CopySheet(ByVal source As Object, ByVal dest As Object,
                         ByVal afterflag As Boolean)
        Try
            If afterflag Then
                source.Copy(After:=dest)
            Else
                source.Copy(Before:=dest)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [DeleteSheet]：Worksheetを削除するメソッド１ </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <param name="sheetname">シート名</param>
    ''' <remarks></remarks>
    Public Sub DeleteSheet(ByVal sheets As Object, ByVal sheetname As String)

        If ExistSheet(sheets, sheetname) Then
            Dim sheet1 = sheets.Item(sheetname)
            sheet1.Delete()
            ReleaseObject(sheet1)
        Else
            Dim msg = "存在しないシートが指定されました。(" & sheetname & ")"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [DeleteSheet]：Worksheetを削除するメソッド２ </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <param name="index">シート番号</param>
    ''' <remarks></remarks>
    Public Sub DeleteSheet(ByVal sheets As Object, ByVal index As Integer)

        Dim count As Integer = sheets.Count
        If count < index Then
            Dim msg = "引数[index] (= " & index.ToString & ") が不適切です。"
            Throw New ArgumentException(msg)
        End If

        Dim sheet1 = sheets.Item(index)
        sheet1.Delete()
        ReleaseObject(sheet1)
    End Sub

    ''' <summary> [MoveSheet]：Worksheetの位置を移動するメソッド </summary>
    ''' <param name="source">移動対象のWorksheet</param>
    ''' <param name="dest">コピー先位置決め用のWorksheet</param>
    ''' <param name="afterflag">True:destの後ろ、False:destの前に移動</param>
    ''' <remarks></remarks>
    Public Sub MoveSheet(ByVal source As Object, ByVal dest As Object,
                         ByVal afterflag As Boolean)
        Try
            If afterflag Then
                source.Move(After:=dest)
            Else
                source.Move(Before:=dest)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [PreviewSheet]：Worksheetのプレビューを表示するメソッド </summary>
    ''' <param name="sheet">Worksheet</param>
    ''' <remarks></remarks>
    Public Sub PreviewSheet(ByVal sheet As Object)
        Try
            sheet.PrintPreview()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetSheetName]：Worksheetにシート名を設定するメソッド </summary>
    ''' <param name="sheet">Worksheet</param>
    ''' <param name="sname">シート名</param>
    ''' <remarks></remarks>
    Public Sub SetSheetName(ByVal sheet As Object, ByVal sname As String)
        Try
            sheet.Name = sname
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    ''' <summary> [VisibleSheet]：Worksheetの表示・非表示を設定するメソッド </summary>
    ''' <param name="sheet">Worksheet</param>
    ''' <param name="flag">True:表示、False:非表示</param>
    ''' <remarks></remarks>
    Public Sub VisibleSheet(ByVal sheet As Object, ByVal flag As Boolean)
        Try
            sheet.Visible = flag
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region "Cells関係"

    ''' <summary> [GetCell]：Cellを取得するメソッド </summary>
    ''' <param name="cells">Cellsオブジェクト</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <returns>指定されたセルを返す</returns>
    ''' <remarks></remarks>
    Public Function GetCell(ByVal cells As Object, ByVal row As Integer, _
                            ByVal col As Integer) As Object
        If row > 0 AndAlso col > 0 Then
            Try
                Return cells.Item(row, col)
            Catch ex As Exception
                Throw
            End Try
        Else
            Dim msg = "[GetCell]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [GetCells]：Cellsを取得するメソッド </summary>
    ''' <param name="sheet">Worksheet</param>
    ''' <returns>Cellsを返す</returns>
    ''' <remarks></remarks>
    Public Function GetCells(ByVal sheet As Object) As Object
        Try
            Return sheet.Cells
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [GetCellValue]：CellのValueを取得するメソッド１ </summary>
    ''' <param name="cell">セル</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCellValue(ByVal cell As Object) As Object
        Try
            Return cell.Value2
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [GetCellValue]：CellのValueを取得するメソッド２ </summary>
    ''' <param name="cells">Cellsオブジェクト</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCellValue(ByVal cells As Object, ByVal row As Integer, _
                                 ByVal col As Integer) As Object

        If row > 0 AndAlso col > 0 Then
            Dim cell1 = GetCell(cells, row, col)
            Dim cellvalue = GetCellValue(cell1)
            ReleaseObject(cell1)
            Return cellvalue
        Else
            Dim msg = "[GetCellValue]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [MergeCells]：Cellを結合するメソッド </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row1">始点セル行番号</param>
    ''' <param name="col1">始点セル列番号</param>
    ''' <param name="row2">終点セル行番号</param>
    ''' <param name="col2">終点セル列番号</param>
    ''' <remarks></remarks>
    Public Sub MergeCells(ByVal vsheet As Object, ByVal row1 As Integer, _
                          ByVal col1 As Integer, ByVal row2 As Integer, _
                          ByVal col2 As Integer)
        Try
            Dim range1 = GetRange(vsheet, row1, col1, row2, col2)
            Try
                _excel.DisplayAlerts = False
                range1.Merge()
                _excel.DisplayAlerts = True
            Catch
                Throw
            Finally
                ReleaseObject(range1)
            End Try

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetCellValue]：CellにValueを記入するメソッド１ </summary>
    ''' <param name="cell"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub SetCellValue(ByVal cell As Object, ByVal value As Object)
        Try
            cell.Value2 = value
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetCellValue]：CellにValueを記入するメソッド２ </summary>
    ''' <param name="cells"></param>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub SetCellValue(ByVal cells As Object, ByVal row As Integer, _
                                ByVal col As Integer, ByVal value As Object)
        If row > 0 AndAlso col > 0 Then
            Dim cell1 = GetCell(cells, row, col)
            SetCellValue(cell1, value)
            ReleaseObject(cell1)
        Else
            Dim msg = "[SetCellValue]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

#End Region

#Region "Range関係"

    ''' <summary> [GetRange]：Rangeを取得するメソッド１ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A1"等</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRange(ByVal vsheet As Object, ByVal srange As String) As Object

        Try
            Return vsheet.Range(srange)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary> [GetRange]：Rangeを取得するメソッド２ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRange(ByVal vsheet As Object, ByVal row As Integer, _
                             ByVal col As Integer) As Object

        If row > 0 AndAlso col > 0 Then
            Try
                Return vsheet.Cells(row, col)
            Catch ex As Exception
                Throw
            End Try
        Else
            Dim msg = "[GetRange]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [GetRange]：Rangeを取得するメソッド３ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row1">セル1行番号</param>
    ''' <param name="col1">セル1列番号</param>
    ''' <param name="row2">セル2行番号</param>
    ''' <param name="col2">セル2列番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRange(ByVal vsheet As Object, ByVal row1 As Integer, _
                             ByVal col1 As Integer, ByVal row2 As Integer, _
                             ByVal col2 As Integer) As Object

        If row1 > 0 AndAlso col1 > 0 AndAlso row2 > 0 AndAlso col2 > 0 Then
            Dim range1 = GetRange(vsheet, row1, col1)
            Dim range2 = GetRange(vsheet, row2, col2)
            Dim range = vsheet.Range(range1, range2)
            ReleaseObject(range1)
            'ReleaseObject(range1)
            ReleaseObject(range2)
            'ReleaseObject(range2)
            Return range
        Else
            Dim msg = "[GetRange]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [GetRangeValue]：Rangeの値を取得するメソッド１ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A1"等</param>
    ''' <returns>指定範囲srangeが不適切な場合には Nothing を返す</returns>
    ''' <remarks></remarks>
    Public Function GetRangeValue(ByVal vsheet As Object, ByVal srange As String) As Object

        Try
            Dim value1 As Object = Nothing
            Dim range1 = GetRange(vsheet, srange)
            Try
                value1 = range1.Value
            Catch
                Throw
            Finally
                ReleaseObject(range1)
            End Try
            Return value1
        Catch ex As Exception
            Throw
        End Try

    End Function

    ''' <summary> [GetRangeValue]：Rangeの値を取得するメソッド２ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <returns>指定セルが不適切な場合には Nothing を返す</returns>
    ''' <remarks></remarks>
    Public Function GetRangeValue(ByVal vsheet As Object, ByVal row As Integer, _
                                  ByVal col As Integer) As Object

        Dim value1 As Object = Nothing
        If row > 0 AndAlso col > 0 Then
            Dim range1 = GetRange(vsheet, row, col)
            value1 = range1.Value
            ReleaseObject(range1)
            Return value1
        Else
            Dim msg = "[GetRangeValue]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [ClearContents]：Rangeのデータだけをクリアするメソッド </summary>
    ''' <param name="vrange"></param>
    ''' <remarks></remarks>
    Public Sub ClearContents(ByVal vrange As Object)
        vrange.ClearContents()
    End Sub

    ''' <summary> [ClearFormats]：Rangeの書式だけをクリアするメソッド </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <remarks></remarks>
    Public Sub ClearFormats(ByVal vrange As Object)
        vrange.ClearFormats()
    End Sub

    ''' <summary> [ClearRange]：Rangeのデータと書式をクリアするメソッド </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <remarks></remarks>
    Public Sub ClearRange(ByVal vrange As Object)
        vrange.Clear()
    End Sub

    ''' <summary> [PutBorder]：Rangeに罫線を引くメソッド </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <param name="vborder">CellBorder</param>
    ''' <remarks></remarks>
    Public Sub PutBorder(ByVal vrange As Object, ByVal vborder As CellBorder)

        Dim border1 As Object = Nothing
        Try
            Dim index As Integer = vborder.Index
            If index < 5 OrElse index > 10 Then
                border1 = vrange.Borders()
            Else
                border1 = vrange.Borders(vborder.Index)
            End If
            With border1
                .LineStyle = vborder.LineStyle
                .Color = ColorTranslator.ToOle(vborder.Color)
                .Weight = vborder.Weight
            End With
        Catch ex As Exception
            Throw
        Finally
            ReleaseObject(border1)
        End Try
    End Sub

    ''' <summary> [SetNumberFormatLocal]：RangeのNumberFormatLocalを設定するメソッド１ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A1"等</param>
    ''' <param name="nfl">NumberFormatLocal</param>
    ''' <remarks></remarks>
    Public Sub SetNumberFormatLocal(ByVal vsheet As Object, ByVal srange As String, _
                                    ByVal nfl As String)
        Try
            Dim range1 = GetRange(vsheet, srange)
            range1.NumberFormatLocal = nfl
            ReleaseObject(range1)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetNumberFormatLocal]：RangeのNumberFormatLocalを設定するメソッド２ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <param name="nfl">NumberFormatLocal</param>
    ''' <remarks></remarks>
    Public Sub SetNumberFormatLocal(ByVal vsheet As Object, ByVal row As Integer, _
                                    ByVal col As Integer, ByVal nfl As String)

        If row > 0 AndAlso col > 0 Then
            Dim range1 = GetRange(vsheet, row, col)
            Try
                range1.NumberFormatLocal = nfl
            Catch ex As Exception
                Throw
            Finally
                ReleaseObject(range1)
            End Try
        Else
            Dim msg = "[SetNumberFormatLocal]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [SetNumberFormatLocal]：RangeのNumberFormatLocalを設定するメソッド３ </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <param name="nfl">NumberFormatLocal</param>
    ''' <remarks></remarks>
    Public Sub SetNumberFormatLocal(ByVal vrange As Object, ByVal nfl As String)
        Try
            vrange.NumberFormatLocal = nfl
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetRange]：Rangeに値を設定するメソッド１ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A1"等</param>
    ''' <param name="ovalue">設定値</param>
    ''' <remarks></remarks>
    Public Sub SetRange(ByVal vsheet As Object, ByVal srange As String, _
                        ByVal ovalue As Object)

        Try
            Dim range1 = GetRange(vsheet, srange)
            range1.Value = ovalue
            ReleaseObject(range1)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetRange]：Rangeに値を設定するメソッド２ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <param name="ovalue">設定値</param>
    ''' <remarks></remarks>
    Public Sub SetRange(ByVal vsheet As Object, ByVal row As Integer, _
                        ByVal col As Integer, ByVal ovalue As Object)

        If row > 0 AndAlso col > 0 Then
            Dim range1 = GetRange(vsheet, row, col)
            range1.Value = ovalue
            ReleaseObject(range1)
        Else
            Dim msg = "[SetRange]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [SetRange]：Rangeに値を設定するメソッド３ </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <param name="ovalue">設定値</param>
    ''' <remarks></remarks>
    Public Sub SetRange(ByVal vrange As Object, ByVal ovalue As Object)
        Try
            vrange.Value = ovalue
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' [SetRangeColor]：Rangeに色を設定するメソッド
    ''' </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <param name="vindex">ColorIndex</param>
    ''' <remarks>2012-12-29追加(Ver1.1.0)</remarks>
    Public Sub SetRangeColor(ByVal vrange As Object, ByVal vindex As Integer)

        Try
            Dim interior As Object = vrange.Interior
            Try
                interior.ColorIndex = vindex
            Catch ex1 As Exception
                Throw
            Finally
                ReleaseObject(interior)
            End Try

        Catch ex2 As Exception
            Throw
        End Try

    End Sub

    ''' <summary> [SetRangeFormula]：RangeのFormulaを設定するメソッド１ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A1"等</param>
    ''' <param name="vformula">Formula</param>
    ''' <remarks></remarks>
    Public Sub SetRangeFormula(ByVal vsheet As Object, ByVal srange As String, _
                               ByVal vformula As String)
        Try
            Dim range1 = GetRange(vsheet, srange)
            range1.Formula = vformula
            ReleaseObject(range1)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetRangeFormula]：RangeのFormulaを設定するメソッド２ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <param name="vformula">Formula</param>
    ''' <remarks></remarks>
    Public Sub SetRangeFormula(ByVal vsheet As Object, ByVal row As Integer, _
                               ByVal col As Integer, ByVal vformula As String)

        If row > 0 AndAlso col > 0 Then
            Dim range1 = GetRange(vsheet, row, col)
            Try
                range1.Formula = vformula
            Catch ex As Exception
                Throw
            Finally
                ReleaseObject(range1)
            End Try
        Else
            Dim msg = "[SetRangeFormula]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [SetRangeFormula]：RangeのFormulaを設定するメソッド３ </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <param name="vformula">Formula</param>
    ''' <remarks></remarks>
    Public Sub SetRangeFormula(ByVal vrange As Object, ByVal vformula As String)
        Try
            vrange.Formula = vformula
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetRangeFormulaR1C1]：RangeのFormulaR1C1を設定するメソッド１ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A1"等</param>
    ''' <param name="vr1c1">FormulaR1C1</param>
    ''' <remarks></remarks>
    Public Sub SetRangeFormulaR1C1(ByVal vsheet As Object, ByVal srange As String, _
                                   ByVal vr1c1 As String)
        Try
            Dim range1 = GetRange(vsheet, srange)
            range1.FormulaR1C1 = vr1c1
            ReleaseObject(range1)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> [SetRangeFormulaR1C1]：RangeのFormulaR1C1を設定するメソッド２ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">セル行番号</param>
    ''' <param name="col">セル列番号</param>
    ''' <param name="vr1c1">FormulaR1C1</param>
    ''' <remarks></remarks>
    Public Sub SetRangeFormulaR1C1(ByVal vsheet As Object, ByVal row As Integer, _
                                   ByVal col As Integer, ByVal vr1c1 As String)
        If row > 0 AndAlso col > 0 Then
            Dim range1 = GetRange(vsheet, row, col)
            Try
                range1.FormulaR1C1 = vr1c1
            Catch ex As Exception
                Throw
            Finally
                ReleaseObject(range1)
            End Try
        Else
            Dim msg = "[SetRangeFormula]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [SetRangeFormulaR1C1]：RangeのFormulaR1C1を設定するメソッド３ </summary>
    ''' <param name="vrange">対象Range</param>
    ''' <param name="vr1c1">FormulaR1C1</param>
    ''' <remarks></remarks>
    Public Sub SetRangeFormulaR1C1(ByVal vrange As Object, ByVal vr1c1 As String)
        Try
            vrange.FormulaR1C1 = vr1c1
        Catch ex As Exception
            Throw
        End Try
    End Sub


#End Region

#Region "Row関係"

    ''' <summary>
    ''' [GetRow]：行を取得するメソッド
    ''' </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">取得する行の番号</param>
    ''' <returns>取得された行を返す</returns>
    ''' <remarks>2012-12-29追加(Ver1.1.0)</remarks>
    Public Function GetRow(ByVal vsheet As Object, ByVal row As Integer) As Object

        If row > 0 Then
            Try
                Dim range1 = GetRange(vsheet, row, 1)
                Try
                    Return range1.EntireRow
                Catch ex1 As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try

            Catch ex2 As Exception
                Throw
            End Try
        Else
            Dim msg = "[GetRow]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [DeleteRow]：行を削除するメソッド </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">削除する行の番号</param>
    ''' <remarks></remarks>
    Public Sub DeleteRow(ByVal vsheet As Object, ByVal row As Integer)

        If row > 0 Then
            Try
                Dim range1 = GetRange(vsheet, row, 1)
                Try
                    Dim entRow As Object = range1.EntireRow
                    Try
                        entRow.Delete()
                    Catch ex1 As Exception
                        Throw
                    Finally
                        ReleaseObject(entRow)
                    End Try

                Catch ex2 As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try
            Catch ex3 As Exception
                Throw
            End Try
        Else
            Dim msg = "[DeleteRow]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [InsertRow]：行を挿入するメソッド </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="row">挿入する行の番号</param>
    ''' <remarks></remarks>
    Public Sub InsertRow(ByVal vsheet As Object, ByVal row As Integer)

        If row > 0 Then
            Try
                Dim range1 = GetRange(vsheet, row, 1)
                Try
                    Dim entRow As Object = range1.EntireRow
                    Try
                        entRow.Insert()
                    Catch ex1 As Exception
                        Throw
                    Finally
                        ReleaseObject(entRow)
                    End Try

                Catch ex2 As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try
            Catch ex3 As Exception
                Throw
            End Try

        Else
            Dim msg = "[InsertRow]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [SetRowHeight]：行高さを設定するメソッド１ </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A6:A9"等</param>
    ''' <param name="height">行高さ設定値</param>
    ''' <remarks></remarks>
    Public Sub SetRowHeight(ByVal vsheet As Object, ByVal srange As String, _
                            ByVal height As Double)

        If height > 0 Then
            Try
                Dim range1 = GetRange(vsheet, srange)
                Try
                    range1.RowHeight = height
                Catch ex As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try

            Catch
                Throw
            End Try
        Else
            Dim msg = "[SetRowHeight]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [SetRowHeight]：行高さを設定するメソッド２ </summary>
    ''' <param name="vsheet">WorkSheet</param>
    ''' <param name="row">行番号</param>
    ''' <param name="height">行高さ設定値</param>
    ''' <remarks></remarks>
    Public Sub SetRowHeight(ByVal vsheet As Object, ByVal row As Integer, _
                            ByVal height As Double)

        If height > 0 AndAlso row > 0 Then
            Try
                Dim range1 = GetRange(vsheet, row, 1)
                Try
                    range1.RowHeight = height
                Catch ex As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try
            Catch
                Throw
            End Try
        Else
            Dim msg = "[SetRowHeight]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

#End Region

#Region "Column関係"

    ''' <summary>
    ''' [GetColumn]：列を取得するメソッド
    ''' </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="col">取得する列の番号</param>
    ''' <returns>取得された列を返す</returns>
    ''' <remarks>2012-12-29追加(Ver1.1.0)</remarks>
    Public Function GetColumn(ByVal vsheet As Object, ByVal col As Integer) As Object

        If col > 0 Then
            Try
                Dim range1 = GetRange(vsheet, 1, col)
                Try
                    Return range1.EntireColumn
                Catch ex1 As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try

            Catch ex2 As Exception
                Throw
            End Try
        Else
            Dim msg = "[GetColumn]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Function

    ''' <summary> [AutoFitColumnWidth]：列幅を自動調整するメソッド </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A:C"等</param>
    ''' <remarks></remarks>
    Public Sub AutoFitColumnWidth(ByVal vsheet As Object, ByVal srange As String)
        Try
            Dim range1 = GetRange(vsheet, srange)
            Dim range2 = range1.EntireColumn

            Try
                range2.AutoFit()
            Catch ex As Exception
                Throw
            Finally
                ReleaseObject(range2)
                ReleaseObject(range1)
            End Try
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' [DeleteColumn]：列を削除するメソッド
    ''' </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="col">削除する列の番号</param>
    ''' <remarks></remarks>
    Public Sub DeleteColumn(ByVal vsheet As Object, ByVal col As Integer)

        If col > 0 Then
            Try
                Dim range1 = GetRange(vsheet, 1, col)
                Try
                    Dim entCol As Object = range1.EntireColumn
                    Try
                        entCol.Delete()
                    Catch ex1 As Exception
                        Throw
                    Finally
                        ReleaseObject(entCol)
                    End Try

                Catch ex2 As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try
            Catch ex3 As Exception
                Throw
            End Try
        Else
            Dim msg = "[DeleteColumn]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary>
    ''' [InsertColumn]：列を挿入するメソッド
    ''' </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="col">挿入する列の番号</param>
    ''' <remarks></remarks>
    Public Sub InsertColumn(ByVal vsheet As Object, ByVal col As Integer)

        If col > 0 Then
            Try
                Dim range1 = GetRange(vsheet, 1, col)
                Try
                    Dim entCol As Object = range1.EntireColumn
                    Try
                        entCol.Insert()
                    Catch ex1 As Exception
                        Throw
                    Finally
                        ReleaseObject(entCol)
                    End Try

                Catch ex2 As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try
            Catch ex3 As Exception
                Throw
            End Try

        Else
            Dim msg = "[InsertColumn]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

    ''' <summary> [SetColumnWidth]：列幅を設定するメソッド </summary>
    ''' <param name="vsheet">Worksheet</param>
    ''' <param name="srange">"A6:A9"等</param>
    ''' <param name="width">列幅設定値</param>
    ''' <remarks></remarks>
    Public Sub SetColumnWidth(ByVal vsheet As Object, ByVal srange As String, _
                              ByVal width As Double)

        If width > 0 Then
            Try
                Dim range1 = GetRange(vsheet, srange)
                Try
                    range1.ColumnWidth = width
                Catch ex As Exception
                    Throw
                Finally
                    ReleaseObject(range1)
                End Try
            Catch ex As Exception
                Throw
            End Try
        Else
            Dim msg = "[SetColumnWidth]メソッドの引数に問題があります。"
            Throw New ArgumentException(msg)
        End If
    End Sub

#End Region

#Region "Privateメソッド"

    ''' <summary> [ExistSheet]：シートの存在を確認するメソッド </summary>
    ''' <param name="sheets">Worksheets</param>
    ''' <param name="sheetname">シート名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ExistSheet(ByVal sheets As Object, ByVal sheetname As String) _
            As Boolean

        Dim result As Boolean = False

        Dim sname = StrConv(sheetname.Trim.ToUpper, VbStrConv.Wide)
        For Each ws In sheets
            If StrConv(ws.Name.Trim.ToUpper, VbStrConv.Wide) = sname Then
                result = True
                ReleaseObject(ws)
                Exit For
            End If
            ReleaseObject(ws)
        Next

        Return result
    End Function

#End Region

End Class
