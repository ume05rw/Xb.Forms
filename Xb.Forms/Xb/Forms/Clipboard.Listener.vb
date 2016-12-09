Option Strict On

Imports System.Windows.Forms
Imports System.Runtime.InteropServices


Public Partial Class Clipboard

    ''' <summary>
    ''' クリップボード監視クラス
    ''' </summary>
    ''' <remarks>
    ''' TODO: Formに対する通知を取得している。Formに依存せず、直接取得出来ないものか。
    ''' 
    ''' クリップボードの内容をリアルタイムに取得するには？
    ''' http://www.atmarkit.co.jp/fdotnet/dotnettips/848cbviewer/cbviewer.html
    ''' </remarks>
    <System.Security.Permissions.PermissionSet( _
        System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
    Public Class Listener

                    Inherits System.Windows.Forms.NativeWindow
        Implements IDisposable

        ''' <summary>
        ''' 値変更イベントの引数クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class ChangeEventArgs
            Inherits EventArgs

            Private _string As String
            Private _html As String
            Private _image As System.Drawing.Image
            Private _datatable As System.Data.DataTable
            Private _format As Xb.Forms.Clipboard.Format


            ''' <summary>
            ''' クリップボード上の値の型
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property Format() As Xb.Forms.Clipboard.Format
                Get
                    Return Me._format
                End Get
            End Property


            ''' <summary>
            ''' 値変更時点のクリップボード上の文字列
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property [String]() As String
                Get
                    Return Me._string
                End Get
            End Property


            ''' <summary>
            ''' 値変更時点のクリップボード上のHTML
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property Html() As String
                Get
                    Return Me._html
                End Get
            End Property


            ''' <summary>
            ''' 値変更時点のクリップボード上の画像
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property Image() As System.Drawing.Image
                Get
                    Return Me._image
                End Get
            End Property


            ''' <summary>
            ''' 値変更時点のクリップボード上のCSVデータ
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property DataTable() As System.Data.DataTable
                Get
                    Return Me._datatable
                End Get
            End Property


            ''' <summary>
            ''' コンストラクタ
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub New()
                Me._format = Xb.Forms.Clipboard.GetFormat()

                Select Case Me._format
                    Case Xb.Forms.Clipboard.Format.String
                        Me._string = Xb.Forms.Clipboard.GetString()
                    Case Xb.Forms.Clipboard.Format.Html
                        Me._string = Xb.Forms.Clipboard.GetString()
                        Me._string = Xb.Forms.Clipboard.GetHtml()
                    Case Xb.Forms.Clipboard.Format.Image
                        Me._image = Xb.Forms.Clipboard.GetImage()
                    Case Xb.Forms.Clipboard.Format.Csv
                        Me._string = Xb.Forms.Clipboard.GetString()
                        Me._datatable = Xb.Forms.Clipboard.GetDataTable()
                End Select
            End Sub

        End Class


        Public Delegate Sub ChangeEventHandler(ByVal sender As Object, ByVal ev As ChangeEventArgs)
        Public Event Changed As ChangeEventHandler


        <DllImport("user32")> _
        Private Shared Function SetClipboardViewer(ByVal hWndNewViewer As IntPtr) As IntPtr
        End Function

        <DllImport("user32")> _
        Private Shared Function ChangeClipboardChain(ByVal hWndRemove As IntPtr, _
                                                    ByVal hWndNewNext As IntPtr) As Boolean
        End Function

        <DllImport("user32")> _
        Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, _
                                            ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        End Function

        Private Const WmDrawclipboard As Integer = &H308
        Private Const WmChangecbchain As Integer = &H30D


        Private _nextHandle As IntPtr
        Private _dummyControlInUiThread As Windows.Forms.Control


        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(Optional ByVal controlUiThread As Windows.Forms.Control = Nothing)

            If (controlUiThread Is Nothing) Then controlUiThread = new Windows.Forms.Control()
            Me._dummyControlInUiThread = controlUiThread

            AssignHandle(Me._dummyControlInUiThread.Handle)
            Me._nextHandle = SetClipboardViewer(Me.Handle)

        End Sub


        ''' <summary>
        ''' ウインドウメッセージ取得時の処理
        ''' </summary>
        ''' <param name="message"></param>
        ''' <remarks>
        ''' </remarks>
        Protected Overloads Overrides Sub WndProc(ByRef message As Message)

            Select Case message.Msg

                Case WmDrawclipboard
                    'ダミーFormオブジェクトをハンドルしているため、本関数は常にUIスレッドで実行される。
                    'UIスレッド取得処理を通さずRaiseする。
                    RaiseEvent Changed(Me, New ChangeEventArgs())

                    If CInt(_nextHandle) <> 0 Then
                        SendMessage(_nextHandle, message.Msg, message.WParam, message.LParam)
                    End If
                    Exit Select

                    ' クリップボード・ビューア・チェーンが更新された
                Case WmChangecbchain
                    If (message.WParam = _nextHandle) Then
                        _nextHandle = message.LParam ' DirectCast(message.LParam, IntPtr)
                    ElseIf (CInt(_nextHandle) <> 0) Then
                        SendMessage(_nextHandle, message.Msg, message.WParam, message.LParam)
                    End If
                    Exit Select

            End Select

            MyBase.WndProc(message)
        End Sub


        Private _disposedValue As Boolean = False        ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me._disposedValue Then
                If disposing Then
                    Xb.Util.Out("Xb.Forms.Clipboard.Listener Disposing...")
                    ' ビューアを解除
                    Dim sts As Boolean = ChangeClipboardChain(Me.Handle, _nextHandle)
                    ReleaseHandle()
                    Me._dummyControlInUiThread.Dispose()
                End If
            End If
            Me._disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region


    End Class

End Class



