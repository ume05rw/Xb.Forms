Option Strict On


Public Partial Class Control

    ''' <summary>
    ''' コントロールが所属するフォームオブジェクトを取得する。
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetForm(ByRef control As Windows.Forms.Control) As Windows.Forms.Form

        If (control Is Nothing) Then
            Throw New ArgumentException("Formオブジェクトが検出出来ません。")
        End If

        If (TypeOf control Is Windows.Forms.Form) Then
            Return DirectCast(control, Windows.Forms.Form)
        Else
            Return Xb.Forms.Control.GetForm(control.Parent)
        End If

    End Function

    ''' <summary>
    ''' フォーム内のコントロールを名前で探して取得する
    ''' </summary>
    ''' <param name="control"></param>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' フォーム内のコントロールを名前で探して取得する
    ''' http://jeanne.wankuma.com/tips/vb.net/form/findcontrol.html
    ''' </remarks>
    Public Shared Function FindControl(ByVal control As Windows.Forms.Control, ByVal name As String) As Windows.Forms.Control

        ' control 内のすべてのコントロールを列挙する
        For Each ctrl As Windows.Forms.Control In control.Controls

            ' コントロール名が合致した場合はそのコントロールのインスタンスを返す
            If (ctrl.Name = name) Then Return ctrl

            ' 列挙したコントロールにコントロールが含まれている場合は再帰呼び出しする
            If (ctrl.HasChildren) Then
                Dim cFindControl As Windows.Forms.Control = Xb.Forms.Control.FindControl(ctrl, name)

                ' 再帰呼び出し先でコントロールが見つかった場合はそのまま返す
                If (Not cFindControl Is Nothing) Then Return cFindControl
            End If
        Next

        Throw New ArgumentException("指定名のコントロールが検出出来ませんでした。")
    End Function

    ''' <summary>
    ''' 指定コントロール配下の全ての子コントロールを取得する。
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAllChildren(ByVal control As Windows.Forms.Control) As List(Of Windows.Forms.Control)
        Dim result As List(Of Windows.Forms.Control) = New List(Of Windows.Forms.Control)()

        ' control 内のすべてのコントロールを列挙する
        For Each ctrl As Windows.Forms.Control In control.Controls
            '直下の子コントロールを追加
            result.Add(ctrl)

            '子コントロールに子があるとき、再帰取得する。
            If (ctrl.Controls IsNot Nothing AndAlso ctrl.Controls.Count > 0) Then
                result.AddRange(Xb.Forms.Control.GetAllChildren(ctrl))
            End If
        Next

        Return result
    End Function


    ''' <summary>
    ''' コントロールにバルーン型ツールティップを追加する。
    ''' </summary>
    ''' <param name="control"></param>
    ''' <param name="message"></param>
    ''' <param name="icon"></param>
    ''' <remarks>
    ''' ToolTip.IsBalloonプロパティを使用する方法
    ''' http://dobon.net/vb/dotnet/control/balloontooltip.html
    ''' </remarks>
    Public Shared Sub SetTooltip(ByVal control As Windows.Forms.Control, _
                                    ByVal message As String, _
                                    Optional ByVal title As String = "", _
                                    Optional ByVal icon As Xb.Forms.Util.DialogIconType = Xb.Forms.Util.DialogIconType.Info)

        Dim toolTip1 As New System.Windows.Forms.ToolTip()

        'ツールチップをバルーンウィンドウとして表示する
        toolTip1.IsBalloon = True

        'ツールチップのタイトル
        toolTip1.ToolTipTitle = title

        'ツールチップに表示するアイコン
        Select Case icon
            Case Xb.Forms.Util.DialogIconType.Info
                toolTip1.ToolTipIcon = Windows.Forms.ToolTipIcon.Info
            Case Xb.Forms.Util.DialogIconType.Error
                toolTip1.ToolTipIcon = Windows.Forms.ToolTipIcon.Error
            Case Xb.Forms.Util.DialogIconType.Warning
                toolTip1.ToolTipIcon = Windows.Forms.ToolTipIcon.Warning
            Case Else
                toolTip1.ToolTipIcon = Windows.Forms.ToolTipIcon.Info
        End Select

        'ツールチップを表示するコントロールと、表示するメッセージ
        toolTip1.SetToolTip(control, message)

    End Sub

End Class



