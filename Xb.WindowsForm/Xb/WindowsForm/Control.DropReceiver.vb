Option Strict On

Imports System.Windows.Forms
Imports System.ComponentModel

Partial Public Class Control

    ''' <summary>
    ''' ドラッグ＆ドロップ制御クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DropReceiver
        Implements IDisposable


        ''' <summary>
        ''' ドラッグ＆ドロップ対象コントロール
        ''' </summary>
        ''' <remarks></remarks>
        Private _control As System.Windows.Forms.Control

        Private _targetType As TargetTypeItems

        ''' <summary>
        ''' 対象データ種
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum TargetTypeItems
            File
            Text
            'Image
        End Enum


#Region "イベント定義"

        ''' <summary>
        ''' ファイル取得イベント定義
        ''' </summary>
        ''' <remarks></remarks>
        Public Class FileRecieveEventArgs
            Inherits EventArgs

            Private _fileNames As String()

            Public ReadOnly Property FileNames() As String()
                Get
                    Return Me._fileNames
                End Get
            End Property

            Public Sub New(ByVal fileNames As String())
                Me._fileNames = fileNames
            End Sub
        End Class
        Public Delegate Sub FileRecieveHandler(ByVal sender As Object, ByVal e As FileRecieveEventArgs)
        Public Event FileRecieved As FileRecieveHandler

        ''' <summary>
        ''' テキスト取得イベント定義
        ''' </summary>
        ''' <remarks></remarks>
        Public Class TextRecieveEventArgs
            Inherits EventArgs

            Private _text As String

            Public ReadOnly Property Text() As String
                Get
                    Return Me._text
                End Get
            End Property

            Public Sub New(ByVal text As String)
                Me._text = text
            End Sub
        End Class
        Public Delegate Sub TextRecieveHandler(ByVal sender As Object, ByVal e As TextRecieveEventArgs)
        Public Event TextRecieved As TextRecieveHandler

        ' ''' <summary>
        ' ''' 画像取得イベント定義
        ' ''' </summary>
        ' ''' <remarks></remarks>
        'Public Class ImageRecieveEventArgs
        '    Inherits EventArgs

        '    Private _image As System.Drawing.Image

        '    Public ReadOnly Property Image() As System.Drawing.Image
        '        Get
        '            Return Me._image
        '        End Get
        '    End Property

        '    Public Sub New(ByVal image As System.Drawing.Image)
        '        Me._image = image
        '    End Sub
        'End Class
        'Public Delegate Sub ImageRecieveHandler(ByVal sender As Object, ByVal e As ImageRecieveEventArgs)
        'Public Event ImageRecieved As ImageRecieveHandler

#End Region


        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(ByVal control As System.Windows.Forms.Control, _
                        ByVal targetType As TargetTypeItems)

            Me._control = control
            Me._targetType = targetType

            Me._control.AllowDrop = True
            AddHandler Me._control.DragEnter, AddressOf Me.DragEnter
            AddHandler Me._control.DragDrop, AddressOf Me.DragDrop

        End Sub


        ''' <summary>
        ''' マウスオーバー時のイベント処理
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)

            Dim detected As Boolean = False
            Select Case Me._targetType
                Case TargetTypeItems.File
                    If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
                        e.Effect = DragDropEffects.Copy
                        detected = True
                    End If
                Case TargetTypeItems.Text
                    If (e.Data.GetDataPresent(DataFormats.Text)) Then
                        e.Effect = DragDropEffects.Copy
                        detected = True
                    End If
                    'Case TargetTypeItems.Image
                    '    If (e.Data.GetDataPresent(DataFormats.Bitmap)) Then
                    '        e.Effect = DragDropEffects.Copy
                    '        detected = True
                    '    End If
                Case Else
                    Throw New InvalidEnumArgumentException("未知の対象データ種です。")
            End Select

            If (Not detected) Then
                e.Effect = DragDropEffects.None
            End If

        End Sub

        ''' <summary>
        ''' ドロップ時のイベント処理
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)

            'Me.writeType(e)

            Select Case Me._targetType
                Case TargetTypeItems.File
                    If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
                        Dim result() As String _
                            = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

                        RaiseEvent FileRecieved(Me, New FileRecieveEventArgs(result))
                    End If
                Case TargetTypeItems.Text
                    If (e.Data.GetDataPresent(DataFormats.Text)) Then
                        Dim result As String = CStr(e.Data.GetData(GetType(String)))
                        RaiseEvent TextRecieved(Me, New TextRecieveEventArgs(result))
                    End If
                    'Case TargetTypeItems.Image
                    '    If (e.Data.GetDataPresent(DataFormats.Bitmap)) Then
                    '        'RaiseEvent ImageRecieved(Me, New ImageRecieveEventArgs(args))
                    '    End If
                Case Else
                    Throw New InvalidEnumArgumentException("未知の対象データ種です。")
            End Select

        End Sub


        ''' <summary>
        ''' ドロップ対象のデータ型候補をコンソールに書き出す。
        ''' </summary>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub writeType(ByVal e As System.Windows.Forms.DragEventArgs)

            Xb.Util.Out("")
            Xb.Util.Out("")

            If (e.Data.GetDataPresent(DataFormats.Bitmap)) Then
                Xb.Util.Out("Bitmap")
            End If
            If (e.Data.GetDataPresent(DataFormats.CommaSeparatedValue)) Then
                Xb.Util.Out("CommaSeparatedValue")
            End If
            If (e.Data.GetDataPresent(DataFormats.Dib)) Then
                Xb.Util.Out("Dib")
            End If
            If (e.Data.GetDataPresent(DataFormats.Dif)) Then
                Xb.Util.Out("Dif")
            End If
            If (e.Data.GetDataPresent(DataFormats.EnhancedMetafile)) Then
                Xb.Util.Out("EnhancedMetafile")
            End If
            If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
                Xb.Util.Out("FileDrop")
            End If
            If (e.Data.GetDataPresent(DataFormats.Html)) Then
                Xb.Util.Out("Html")
            End If
            If (e.Data.GetDataPresent(DataFormats.Locale)) Then
                Xb.Util.Out("Locale")
            End If
            If (e.Data.GetDataPresent(DataFormats.MetafilePict)) Then
                Xb.Util.Out("MetafilePict")
            End If
            If (e.Data.GetDataPresent(DataFormats.OemText)) Then
                Xb.Util.Out("OemText")
            End If
            If (e.Data.GetDataPresent(DataFormats.Palette)) Then
                Xb.Util.Out("Palette")
            End If
            If (e.Data.GetDataPresent(DataFormats.PenData)) Then
                Xb.Util.Out("PenData")
            End If
            If (e.Data.GetDataPresent(DataFormats.Riff)) Then
                Xb.Util.Out("Riff")
            End If
            If (e.Data.GetDataPresent(DataFormats.Rtf)) Then
                Xb.Util.Out("Rtf")
            End If
            If (e.Data.GetDataPresent(DataFormats.Serializable)) Then
                Xb.Util.Out("Serializable")
            End If
            If (e.Data.GetDataPresent(DataFormats.StringFormat)) Then
                Xb.Util.Out("StringFormat")
            End If
            If (e.Data.GetDataPresent(DataFormats.SymbolicLink)) Then
                Xb.Util.Out("SymbolicLink")
            End If
            If (e.Data.GetDataPresent(DataFormats.Text)) Then
                Xb.Util.Out("Text")
            End If
            If (e.Data.GetDataPresent(DataFormats.Tiff)) Then
                Xb.Util.Out("Tiff")
            End If
            If (e.Data.GetDataPresent(DataFormats.UnicodeText)) Then
                Xb.Util.Out("UnicodeText")
            End If
            If (e.Data.GetDataPresent(DataFormats.WaveAudio)) Then
                Xb.Util.Out("WaveAudio")
            End If

        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then

                    If (Me._control IsNot Nothing) Then
                        RemoveHandler Me._control.DragEnter, AddressOf Me.DragEnter
                        RemoveHandler Me._control.DragDrop, AddressOf Me.DragDrop
                        Me._control = Nothing
                    End If
                End If
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: 上の Dispose(ByVal disposing As Boolean) にアンマネージ リソースを解放するコードがある場合にのみ、Finalize() をオーバーライドします。
        'Protected Overrides Sub Finalize()
        '    ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Class



