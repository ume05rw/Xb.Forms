Option Strict On


''' <summary>
''' クリップボード関連関数群
''' </summary>
''' <remarks>
''' Sharedメソッドが動作しないときは、動作中のスレッドがSTAでない可能性がある。
''' クリップボードのデータの取得、設定がうまくいかないときは
''' http://dobon.net/vb/dotnet/system/nostaclipboard.html
''' 
''' </remarks>
Public Partial Class Clipboard

            ''' <summary>
    ''' クリップボード上の値の型を示す。
    ''' </summary>
    ''' <remarks>
    ''' 本クラスで扱えるもののみを列挙している。
    ''' </remarks>
    Public Enum Format
        ''' <summary>
        ''' 文字列型
        ''' </summary>
        ''' <remarks></remarks>
        [String]

        ''' <summary>
        ''' HTML形式文字列
        ''' </summary>
        ''' <remarks>
        ''' 値はGetString(文字列)で取得する。
        '''   -> 文字列取得後のHTMLパースが自力で出来ないのでHtmlAgilityPack依存になってしまう。
        '''   -> クラス単体切り出しが難しくなるため、棚上げ。
        ''' </remarks>
        Html

        ''' <summary>
        ''' 画像
        ''' </summary>
        ''' <remarks></remarks>
        Image

        ''' <summary>
        ''' カンマ区切り文字列
        ''' </summary>
        ''' <remarks></remarks>
        Csv

        ''' <summary>
        ''' 値が存在しない
        ''' </summary>
        ''' <remarks></remarks>
        NoData
    End Enum


    ''' <summary>
    ''' 現在クリップボードにある値の型を返す。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFormat() As Xb.Forms.Clipboard.Format

        Dim result As Xb.Forms.Clipboard.Format = Xb.Forms.Clipboard.Format.NoData, _
            data As Windows.Forms.IDataObject = System.Windows.Forms.Clipboard.GetDataObject()

        '値が存在しないとき、それを示す型を返す。
        If (data Is Nothing) Then Return Xb.Forms.Clipboard.Format.NoData

        For Each format As String In data.GetFormats()
            Select Case format
                Case Windows.Forms.DataFormats.Text, _
                    Windows.Forms.DataFormats.UnicodeText

                    '既にHTML, CSVが選択されているとき、上書きしない。
                    If (result = Xb.Forms.Clipboard.Format.Html _
                        OrElse result = Xb.Forms.Clipboard.Format.Csv) Then Exit Select

                    result = Xb.Forms.Clipboard.Format.String

                Case System.Windows.Forms.DataFormats.CommaSeparatedValue

                    result = Xb.Forms.Clipboard.Format.Csv

                Case System.Windows.Forms.DataFormats.Html

                    result = Xb.Forms.Clipboard.Format.Html

                Case Windows.Forms.DataFormats.Bitmap, _
                    Windows.Forms.DataFormats.Dib, _
                    Windows.Forms.DataFormats.MetafilePict, _
                    Windows.Forms.DataFormats.Tiff

                    result = Xb.Forms.Clipboard.Format.Image
                Case Else
                    '何もしない。
            End Select
        Next

        Return result

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値が文字列か否かを調べる。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' ※注意※ HTML, CSV でも True を返す。
    ''' </remarks>
    Public Shared Function IsString() As Boolean

        Dim format As Xb.Forms.Clipboard.Format = GetFormat()
        Return (format = Xb.Forms.Clipboard.Format.String _
                OrElse format = Xb.Forms.Clipboard.Format.Html _
                OrElse format = Xb.Forms.Clipboard.Format.Csv)

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値がHTMLか否かを調べる。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsHtml() As Boolean

        Return (GetFormat() = Xb.Forms.Clipboard.Format.Html)

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値が画像か否かを調べる。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsImage() As Boolean

        Return (GetFormat() = Xb.Forms.Clipboard.Format.Image)

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値がCSVか否かを調べる。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsCsv() As Boolean

        Return (GetFormat() = Xb.Forms.Clipboard.Format.Csv)

    End Function


    ''' <summary>
    ''' クリップボード上の値を削除する。
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Clear()

        Windows.Forms.Clipboard.SetDataObject(New Windows.Forms.DataObject())

    End Sub


    ''' <summary>
    ''' クリップボードに文字列をセットする。
    ''' </summary>
    ''' <param name="text"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetString(ByVal text As String)

        System.Windows.Forms.Clipboard.SetText(text)

    End Sub


    ''' <summary>
    ''' クリップボードに画像をセットする。
    ''' </summary>
    ''' <param name="image"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetImage(ByVal image As System.Drawing.Image)

        System.Windows.Forms.Clipboard.SetImage(image)

    End Sub


    ''' <summary>
    ''' クリップボードにスプレッドシートデータをセットする。
    ''' </summary>
    ''' <param name="datatable"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetDataTable(ByVal datatable As DataTable)

        Dim csvText As String = Xb.File.Csv.GetCsvText(datatable), _
            bytes() As Byte, _
            memoryStream As IO.MemoryStream, _
            dataObject As System.Windows.Forms.DataObject

        bytes = System.Text.Encoding.Default.GetBytes(csvText)
        memoryStream = New IO.MemoryStream(bytes)
        dataObject = New System.Windows.Forms.DataObject(Windows.Forms.DataFormats.CommaSeparatedValue, memoryStream)
        System.Windows.Forms.Clipboard.SetDataObject(dataObject)

    End Sub


    ''' <summary>
    ''' 現在クリップボード上にある値を文字列型として取得する。
    ''' </summary>
    ''' <remarks>
    ''' 取得出来ない場合、Nothingを返す。
    ''' </remarks>
    Public Shared Function GetString() As String

        Dim result As String = System.Windows.Forms.Clipboard.GetText()
        If (result = String.Empty) Then Return Nothing
        Return result

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値をHTML型として、HTML文字列を取得する。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Clipboard.GetData(DataFormats.Html)で文字化け 
    ''' http://gushwell.ldblog.jp/archives/51974858.html
    ''' </remarks>
    Public Shared Function GetHtml() As String

        Dim data As System.Windows.Forms.IDataObject = System.Windows.Forms.Clipboard.GetDataObject(), _
            memoryStream As System.IO.MemoryStream, _
            result As String = "", _
            bytes() As Byte

        'HTMLデータでないとき、空文字列を返す。
        If (Not data.GetDataPresent(Windows.Forms.DataFormats.Html)) Then Return ""

        'memoryStream = DirectCast(data.GetData(Windows.Forms.DataFormats.Html), IO.MemoryStream)
        memoryStream = DirectCast(data.GetData("Html Format"), IO.MemoryStream)
        bytes = memoryStream.ToArray()
        result = Xb.Str.GetEncode(bytes, True).GetString(bytes)

        '注意：以下では文字化けしてしまう。
        'result = CStr(data.GetData(Windows.Forms.DataFormats.Html))

        Return result

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値を画像型として取得する。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 取得出来ない場合、Nothingを返す。
    ''' </remarks>
    Public Shared Function GetImage() As System.Drawing.Image

        Dim result As System.Drawing.Image = System.Windows.Forms.Clipboard.GetImage()
        Return result

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値をスプレッドシートデータとして、DataTable型で取得する。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 取得出来ない場合、Nothingを返す。
    ''' </remarks>
    Public Shared Function GetDataTable() As DataTable

        Dim dataObject As System.Windows.Forms.IDataObject, _
            memoryStream As System.IO.MemoryStream, _
            reader As System.IO.StreamReader, _
            csvText, lines() As String, _
            commaCount, tmpCount As Integer

        dataObject = System.Windows.Forms.Clipboard.GetDataObject()
        If (dataObject Is Nothing) Then Return Nothing

        memoryStream = DirectCast(dataObject.GetData(Windows.Forms.DataFormats.CommaSeparatedValue), _
                                    System.IO.MemoryStream)

        If (memoryStream Is Nothing) Then Return Nothing

        reader = New System.IO.StreamReader(memoryStream, System.Text.Encoding.Default)
        csvText = reader.ReadToEnd()

        csvText = csvText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Trim(CChar(vbLf))
        lines = csvText.Split(CChar(vbLf))

        'CSV形式に整形する。
        'カンマの最大個数を調べる。
        commaCount = 0
        For i As Integer = 0 To lines.Length - 1
            tmpCount = (lines(i).Length - lines(i).Replace(",", "").Length)
            If (commaCount < tmpCount) Then commaCount = tmpCount
        Next

        'カンマ個数が最大個数に満たない行があるとき、行末尾に追加する。
        For i As Integer = 0 To lines.Length - 1
            tmpCount = (lines(i).Length - lines(i).Replace(",", "").Length)
            If (commaCount > tmpCount) Then
                lines(i) &= New String(","c, commaCount - tmpCount)
            End If
        Next

        csvText = String.Join(vbLf, lines).Trim(CChar(vbLf))

        reader.Dispose()

        Return Xb.File.Csv.GetDataTableByText(csvText)

    End Function


    ''' <summary>
    ''' 現在クリップボード上にある値を、DataGridViewのカレントセル以下に張り付ける。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function PasteToDataGridView(ByRef view As System.Windows.Forms.DataGridView) As Boolean

        If (view Is Nothing OrElse view.CurrentCell Is Nothing) Then Return False

        Dim rowIndex, columnIndex As Integer, _
            table As DataTable = Xb.File.Csv.GetDataTableByText(System.Windows.Forms.Clipboard.GetText())

        If (table Is Nothing) Then Return False

        rowIndex = view.CurrentCell.RowIndex
        columnIndex = view.CurrentCell.ColumnIndex

        For i As Integer = 0 To table.Rows.Count - 1
            If (rowIndex + i > view.RowCount - 1) Then Exit For
            For j As Integer = 0 To table.Columns.Count - 1
                If (columnIndex + j > view.ColumnCount - 1) Then Exit For
                Try
                    view.Rows.Item(rowIndex + i).Cells(columnIndex + j).Value = table.Rows.Item(i).Item(j).ToString()
                Catch ex As Exception
                End Try
            Next
        Next

        Return True

    End Function



    '以下、厳密過ぎて用途が無さそうなので使用しない。
    '''' <summary>
    '''' クリップボード値の実際の形式
    '''' </summary>
    '''' <remarks>
    '''' System.Windows.Forms.DataFormats の値候補
    '''' </remarks>
    'Public Enum InnerFormat
    '    ''' <summary>
    '    ''' Microsoft Windows ビットマップ データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' ビットマップはメモリ内のビットの配列としてのコンピューター グラフィックを表し、
    '    ''' ビットはイメージのそれぞれのピクセルの属性を表します。
    '    ''' </remarks>
    '    Bitmap

    '    ''' <summary>
    '    ''' CSV
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' この形式は、スプレッドシートで一般的に使用される基本的な交換形式です。
    '    ''' Grid - Excel のやりとり等に使用。
    '    ''' </remarks>
    '    CommaSeparatedValue

    '    ''' <summary>
    '    ''' デバイスに依存しないビットマップ
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' DIB は、あるアプリケーションで作成されたビットマップ化グラフィックスを、
    '    ''' 他のアプリケーションでも、作成元のアプリケーションと同じように読み込みや
    '    ''' 表示をできるようにするためのファイル形式です。
    '    ''' </remarks>
    '    Dib

    '    ''' <summary>
    '    ''' Windows DIF (Data Interchange Format) データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' DIF は ASCII コードで記述され、データベースやスプレッドシートのような
    '    ''' ドキュメントを複数のプログラムで使用したり交換したりできるように構成されている
    '    ''' フォーマットです。
    '    ''' </remarks>
    '    Dif

    '    ''' <summary>
    '    ''' Windows 拡張メタファイル形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' 拡張メタファイル形式は、イメージをピクセルではなくグラフィックス オブジェクトとして
    '    ''' 格納する Windows ファイルです。 サイズを変更すると、メタファイルはビットマップよりも
    '    ''' 画質の良いイメージを保存します。 
    '    ''' </remarks>
    '    EnhancedMetafile

    '    ''' <summary>
    '    ''' Windows ファイル ドロップ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' ドラッグ アンド ドロップ操作時にシェルのファイル ドラッグと対話するには、
    '    ''' この形式を使用します。
    '    ''' </remarks>
    '    FileDrop

    '    ''' <summary>
    '    ''' HTML データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    Html

    '    ''' <summary>
    '    ''' Windows ロケール (カルチャ) データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    Locale

    '    ''' <summary>
    '    ''' Windows メタファイル画像データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' メタファイル画像形式は、イメージをピクセルではなくグラフィックス オブジェクトとして
    '    ''' 格納する Windows ファイルです。 サイズを変更すると、メタファイルはビットマップよりも
    '    ''' 画質の良いイメージを保存します。 
    '    ''' </remarks>
    '    MetafilePict

    '    ''' <summary>
    '    ''' 標準の Windows OEM テキスト データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    OemText

    '    ''' <summary>
    '    ''' Windows パレット データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    Palette

    '    ''' <summary>
    '    ''' Windows ペン データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' Windows ペン データは、手書きソフトウェア用のペン ストロークで構成されます。
    '    ''' </remarks>
    '    PenData

    '    ''' <summary>
    '    ''' リソース交換ファイル形式 (RIFF) のオーディオ データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' RIFF は汎用性のある仕様で、さまざまな種類のマルチメディア ファイルの標準形式を
    '    ''' 定義する場合に使用されます。
    '    ''' </remarks>
    '    Riff

    '    ''' <summary>
    '    ''' リッチ テキスト形式 (RTF) データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' RTF は、書式設定されたテキスト ドキュメントをアプリケーション間で交換するために
    '    ''' 使用される、Document Content Architecture に基づく形式です。
    '    ''' </remarks>
    '    Rtf

    '    ''' <summary>
    '    ''' 任意の種類のシリアル化可能なデータ オブジェクトをカプセル化するデータ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    Serializable

    '    ''' <summary>
    '    ''' 共通言語ランタイム (CLR) 文字列クラス データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    StringFormat

    '    ''' <summary>
    '    ''' Windows シンボリック リンク データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' シンボリック リンクはディスク ディレクトリ エントリの 1 つで、ファイルの
    '    ''' ディレクトリ エントリを指定しますが、実際には別のディレクトリのファイルを参照しています。 
    '    ''' シンボリック リンクは、エイリアス、ショートカット、ソフト リンク、またはシムリンクと
    '    ''' 呼ばれることもあります。 
    '    ''' </remarks>
    '    SymbolicLink

    '    ''' <summary>
    '    ''' ANSI テキスト データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    Text

    '    ''' <summary>
    '    ''' Tagged Image File Format (TIFF) データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' TIFF は、一般にグレースケール グラフィックス イメージのスキャン、保存、および交換に
    '    ''' 使用される標準ファイル形式です。
    '    ''' </remarks>
    '    Tiff

    '    ''' <summary>
    '    ''' Unicode テキスト データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このデータ形式は、16 ビット文字でエンコードされた Unicode を表し、UTF-16 または 
    '    ''' UCS-2 とも呼ばれます。
    '    ''' </remarks>
    '    UnicodeText

    '    ''' <summary>
    '    ''' WAVE オーディオ データ形式
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' このフィールドは、使用可能なデータ形式を表すために IDataObject インターフェイスを
    '    ''' 実装する DataObject クラスとその他のクラスによって使用されます。
    '    ''' </remarks>
    '    WaveAudio
    'End Enum


    '''' <summary>
    '''' 現在クリップボードにあるデータの、解釈可能な型を全て取得する。
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function GetInnerFormats() As List(Of Xb.Forms.Clipboard.InnerFormat)
    '    Dim formats As List(Of Xb.Forms.Clipboard.InnerFormat) = New List(Of Xb.Forms.Clipboard.InnerFormat)()

    '    Dim data As Windows.Forms.IDataObject = System.Windows.Forms.Clipboard.GetDataObject()


    '    If (Not data Is Nothing) Then
    '        For Each format As String In data.GetFormats()
    '            Select Case format
    '                Case System.Windows.Forms.DataFormats.Bitmap
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Bitmap)
    '                Case System.Windows.Forms.DataFormats.CommaSeparatedValue
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.CommaSeparatedValue)
    '                Case System.Windows.Forms.DataFormats.Dib
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Dib)
    '                Case System.Windows.Forms.DataFormats.Dif
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Dif)
    '                Case System.Windows.Forms.DataFormats.EnhancedMetafile
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.EnhancedMetafile)
    '                Case System.Windows.Forms.DataFormats.FileDrop
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.FileDrop)
    '                Case System.Windows.Forms.DataFormats.Html
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Html)
    '                Case System.Windows.Forms.DataFormats.Locale
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Locale)
    '                Case System.Windows.Forms.DataFormats.MetafilePict
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.MetafilePict)
    '                Case System.Windows.Forms.DataFormats.OemText
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.OemText)
    '                Case System.Windows.Forms.DataFormats.Palette
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Palette)
    '                Case System.Windows.Forms.DataFormats.PenData
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.PenData)
    '                Case System.Windows.Forms.DataFormats.Riff
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Riff)
    '                Case System.Windows.Forms.DataFormats.Rtf
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Rtf)
    '                Case System.Windows.Forms.DataFormats.Serializable
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Serializable)
    '                Case System.Windows.Forms.DataFormats.StringFormat
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.StringFormat)
    '                Case System.Windows.Forms.DataFormats.SymbolicLink
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.SymbolicLink)
    '                Case System.Windows.Forms.DataFormats.Text
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Text)
    '                Case System.Windows.Forms.DataFormats.Tiff
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.Tiff)
    '                Case System.Windows.Forms.DataFormats.UnicodeText
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.UnicodeText)
    '                Case System.Windows.Forms.DataFormats.WaveAudio
    '                    formats.Add(Xb.Forms.Clipboard.InnerFormat.WaveAudio)
    '                Case Else
    '                    '何もしない
    '            End Select
    '        Next
    '    End If

    '    Return formats

    'End Function

End Class


