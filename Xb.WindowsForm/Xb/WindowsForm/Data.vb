Option Strict On

Public Partial Class Data

    ''' <summary>
    ''' DataGridViewの表示内容をCSVフォーマットテキストに変換する。
    ''' </summary>
    ''' <param name="dataGridView"></param>
    ''' <param name="linefeed"></param>
    ''' <param name="isOutputHeader"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCsvText(ByRef dataGridView As Windows.Forms.DataGridView, _
                                        Optional ByVal linefeed As Xb.Str.LinefeedType = Xb.Str.LinefeedType.Lf, _
                                        Optional ByVal isOutputHeader As Boolean = True, _
                                        Optional ByVal isQuote As Boolean = True) As String

        Dim lfString As String, _
            builder As System.Text.StringBuilder, _
            maxColumnCount, maxRowCount, i, j As Integer

        '渡されたDataGridViewデータ存在チェック
        If (dataGridView Is Nothing) Then
            Xb.Util.Out("Xb.File.Csv.GetCsvText: DataTableが検出できません。")
            Throw New ArgumentException("DataTableが検出できません。")
        End If

        lfString = Xb.Str.GetLinefeed(linefeed)
        maxColumnCount = dataGridView.ColumnCount
        maxRowCount = dataGridView.RowCount
        builder = New System.Text.StringBuilder()

        For i = 0 To maxRowCount - 1
            If ((i = 0) And (isOutputHeader)) Then
                '一行目にタイトル行を出力する
                For j = 0 To maxColumnCount - 1
                    builder.Append(IIf(j = 0, "", ",").ToString())
                    builder.Append(If(isQuote, _
                                        Xb.Str.CsvDquote(dataGridView.Columns(j).HeaderText), _
                                        dataGridView.Columns(j).HeaderText))
                Next
                builder.Append(lfString)
            End If

            For j = 0 To maxColumnCount - 1
                builder.Append(IIf(j = 0, "", ",").ToString())
                builder.Append(If(isQuote, _
                                    Xb.Str.CsvDquote(dataGridView.Item(j, i).Value.ToString()), _
                                    dataGridView.Item(j, i).Value.ToString()))
            Next
            builder.Append(lfString)
        Next

        Return builder.ToString()

    End Function
    
    ''' <summary>
    ''' DataGridViewの内容をCSVファイルに書き出す。
    ''' </summary>
    ''' <param name="dataGridView"></param>
    ''' <param name="encode"></param>
    ''' <param name="fileName"></param>
    ''' <param name="isOutputHeader"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' エンコードを特定したい場合に使用。SJIS限定のときは引数違いの同名関数を使用する。
    ''' グリッドのDataSourceバインド有無に関わらず、CSVファイルを生成出来る。
    ''' 
    ''' ※注意※ DataGridViewへのアクセスが遅いので、なるべくDataTableからCSV書き出しをすること。
    ''' 
    ''' </remarks>
    Public Shared Function CreateCsvFile(ByRef dataGridView As Windows.Forms.DataGridView, _
                                                ByVal encode As System.Text.Encoding, _
                                                Optional ByVal linefeed As Xb.Str.LinefeedType = Xb.Str.LinefeedType.Lf, _
                                                Optional ByVal fileName As String = "list.csv", _
                                                Optional ByVal isOutputHeader As Boolean = True, _
                                                Optional ByVal isQuote As Boolean = True) As Boolean

        Dim writer As IO.StreamWriter, _
            csvtext As String

        '渡されたDataGridViewデータ存在チェック
        If (dataGridView Is Nothing) Then Return False

        'ファイル名を絶対パスに整形する。
        Try
            fileName = Xb.App.Path.GetAbsPath(fileName)
        Catch ex As Exception
            Return False
        End Try

        Try
            csvtext = GetCsvText(dataGridView, linefeed, isOutputHeader, isQuote)
            writer = New IO.StreamWriter(fileName, False, encode)
            writer.Write(csvtext)
            writer.Close()
        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function


    ''' <summary>
    ''' DataTableの内容をCSVファイルに書き出す。
    ''' </summary>
    ''' <param name="dataGridView"></param>
    ''' <param name="fileName"></param>
    ''' <param name="isOutputHeader"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateCsvFile(ByRef dataGridView As Windows.Forms.DataGridView, _
                                                Optional ByVal fileName As String = "list.csv", _
                                                Optional ByVal isOutputHeader As Boolean = True) As Boolean

        Return CreateCsvFile(dataGridView, _
                            System.Text.Encoding.GetEncoding("Shift_JIS"), _
                            Xb.Str.LinefeedType.CrLf, _
                            fileName, _
                            isOutputHeader)
    End Function

End Class



