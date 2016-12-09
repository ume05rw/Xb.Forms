Option Strict On

Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Partial Class Util

    ''' <summary>
    ''' ダイアログのアイコン区分
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum DialogIconType
        Info
        [Error]
        Warning
    End Enum

    ''' <summary>
    ''' OKボタンのみのメッセージダイアログを表示する。
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="iconType"></param>
    ''' <remarks></remarks>
    Public Shared Sub Alert(ByVal message As String, _
                            Optional ByVal iconType As DialogIconType = DialogIconType.Error)

        Dim icon As MessageBoxIcon
        Select Case iconType
            Case DialogIconType.Info
                icon = MessageBoxIcon.Information
            Case DialogIconType.Error
                icon = MessageBoxIcon.Error
            Case DialogIconType.Warning
                icon = MessageBoxIcon.Warning
            Case Else
                Throw New ArgumentException("認識出来ないアイコン区分です。")
        End Select

        MessageBox.Show(message, _
                        My.Application.Info.ProductName, _
                        MessageBoxButtons.OK, _
                        icon)

    End Sub

        ''' <summary>
    ''' エラー表示時と併せて例外データをログ出力する。
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="ex"></param>
    ''' <remarks></remarks>
    Public Shared Sub Alert(ByVal message As String, ByVal ex As Exception)
        Dim errText As String = System.String.Join(vbCrLf, Xb.Util.GetErrorString(ex))
        Xb.Util.Out(message & vbCrLf & errText)
        Xb.WindowsForm.Util.Alert(message _
                             & vbCrLf _
                             & vbCrLf _
                             & If(errText.Length > 200, errText.Substring(0, 200) _
                             & "....", errText))

    End Sub

        ''' <summary>
    ''' エラー表示時と併せてモデルエラーデータをログ出力する。
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="errors"></param>
    ''' <remarks></remarks>
    Public Shared Sub Alert(ByVal message As String, ByRef errors As Xb.Db.Model.Error())

        Dim errText, msgText, tmpName, tmpValue As String

        errText = ""
        msgText = ""
        msgText = ""
        For Each err As Xb.Db.Model.Error In errors
            errText &= vbCrLf & err.Message
            msgText &= vbCrLf & err.Message

            tmpName = err.Name.Replace(vbCr, String.Empty) _
                            .Replace(vbLf, String.Empty) _
                            .Replace("-", String.Empty) _
                            .Trim()
            tmpValue = err.Value.Replace(vbCr, String.Empty) _
                            .Replace(vbLf, String.Empty) _
                            .Replace("-", String.Empty) _
                            .Trim()

            If (Not String.IsNullOrEmpty(tmpName) _
               OrElse Not String.IsNullOrEmpty(tmpValue)) Then

                errText &= " ： " _
                    & If(String.IsNullOrEmpty(tmpName), "", " Column = " & tmpName) _
                    & If(String.IsNullOrEmpty(tmpValue), "", " Value = " & tmpValue)

                msgText &= If(String.IsNullOrEmpty(tmpName), "", " ： " & tmpName)
            End If
        Next

        Xb.Util.Out(message & vbCrLf & errText)
        Xb.WindowsForm.Util.Alert(message & vbCrLf & vbCrLf & If(msgText.Length > 200, _
                                                             msgText.Substring(0, 200) & "....", _
                                                             msgText))

    End Sub

    ''' <summary>
    ''' OK/キャンセルボタンのメッセージダイアログを表示する。
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="iconType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Confirm(ByVal message As String, _
                                   Optional ByVal iconType As DialogIconType = DialogIconType.Info) As Boolean

        Dim icon As MessageBoxIcon
        Select Case iconType
            Case DialogIconType.Info
                icon = MessageBoxIcon.Information
            Case DialogIconType.Error
                icon = MessageBoxIcon.Error
            Case DialogIconType.Warning
                icon = MessageBoxIcon.Warning
            Case Else
                Throw New ArgumentException("認識出来ないアイコン区分です。")
        End Select

        Return (MessageBox.Show(message, _
                               My.Application.Info.ProductName, _
                               MessageBoxButtons.OKCancel, _
                               icon) = DialogResult.OK)

    End Function

    Private Shared IndicatorThread As Xb.App.Thread.Executer = Nothing
    Public Shared Sub ShowWaitIndicator()

        If (Util.IndicatorThread Is Nothing) Then
            Util.IndicatorThread = New Xb.App.Thread.Executer(System.Type.GetType("App+WaitingIndicator"))
        End If

        Util.IndicatorThread.Execute("Show")

    End Sub

    Public Shared Sub HideWaitIndicator()

        If (Util.IndicatorThread Is Nothing) Then
            Return
        End If

        Util.IndicatorThread.Execute("Hide")

    End Sub


    ''' <summary>
    ''' 指定ハンドルのスクロール位置を取得する。
    ''' </summary>
    ''' <param name="hWnd"></param>
    ''' <param name="nBar"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Private Shared Function GetScrollPos(hWnd As Integer, nBar As Integer) As Integer
    End Function

    ''' <summary>
    ''' 指定ハンドルのスクロール位置をセットする。
    ''' </summary>
    ''' <param name="hWnd"></param>
    ''' <param name="nBar"></param>
    ''' <param name="nPos"></param>
    ''' <param name="bRedraw"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll")> _
    Private Shared Function SetScrollPos(hWnd As IntPtr, nBar As Integer, nPos As Integer, bRedraw As Boolean) As Integer
    End Function

    Private Const SB_HORZ As Integer = &H0
    Private Const SB_VERT As Integer = &H1


    ''' <summary>
    ''' スクロール可能なコントロールで、スクロール位置を取得する。
    ''' </summary>
    ''' <param name="handle"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Treeview のスクロール位置を維持します。
    ''' http://ja.softuses.com/78601
    ''' </remarks>
    Public Shared Function GetScrollPosition(ByVal handle As IntPtr) As Drawing.Point
        Return New Drawing.Point(GetScrollPos(CInt(handle), SB_HORZ), _
                                 GetScrollPos(CInt(handle), SB_VERT))
    End Function


    ''' <summary>
    ''' スクロール可能コントロールのスクロール位置をセットする。
    ''' </summary>
    ''' <param name="handle"></param>
    ''' <param name="point"></param>
    ''' <remarks>
    ''' Treeview のスクロール位置を維持します。
    ''' http://ja.softuses.com/78601
    ''' </remarks>
    Public Shared Sub SetScrollPosition(ByVal handle As IntPtr, ByVal point As Drawing.Point)
        SetScrollPos(handle, SB_HORZ, point.X, True)
        SetScrollPos(handle, SB_VERT, point.Y, True)
    End Sub


    ''' <summary>
    ''' コントロールが所属するFormオブジェクトを取得する。
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetForm(ByVal control As Windows.Forms.Control) As Windows.Forms.Form
        If (control Is Nothing) Then Return Nothing

        If (TypeOf control Is Windows.Forms.Form) Then
            Return CType(control, Windows.Forms.Form)
        Else
            Return Xb.WindowsForm.Util.GetForm(control.Parent)
        End If

    End Function


    ''' <summary>
    ''' コントロールの所属フォームを最大化表示する。
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks>
    ''' 方法を忘れそうなので、切り出して置いておく。
    ''' SeizableToolWindowが最大化すると、タスクバーの上に表示される、らしい。
    ''' </remarks>
    Public Shared Sub SetFullScreen(ByVal control As Windows.Forms.Control)
        If (control Is Nothing) Then
            Xb.Util.Out("Xb.Ui.SetFullScreen: 渡し値コントロールが検出出来ません。")
            Throw New ArgumentException("渡し値コントロールが検出出来ません。")
        End If

        Dim parentForm As Windows.Forms.Form = GetForm(control)

        '↓順序に注意。
        parentForm.FormBorderStyle = Windows.Forms.FormBorderStyle.SizableToolWindow
        parentForm.WindowState = Windows.Forms.FormWindowState.Maximized
        parentForm.FormBorderStyle = Windows.Forms.FormBorderStyle.None
    End Sub


    
    ''' <summary>
    ''' スリープせずに指定秒数分待機する。
    ''' </summary>
    ''' <param name="second"></param>
    ''' <remarks></remarks>
    Public Shared Sub Wait(ByVal second As Integer)

        Dim watch As Stopwatch = New Stopwatch()
        Dim millisec As Integer = second * 1000
        watch.Start()
        
        Do
            System.Windows.Forms.Application.DoEvents()
        Loop While (watch.ElapsedMilliseconds < millisec)

    End Sub

End Class
