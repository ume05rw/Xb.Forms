Option Strict On


''' <summary>
''' UIパーツの位置・サイズ情報保持クラス
''' </summary>
''' <remarks></remarks>
Public Class ViewParams

    Public X As Integer
    Public Y As Integer
    Public Width As Integer
    Public Height As Integer

    Friend ReadOnly AddX As Integer
    Friend ReadOnly AddY As Integer
    Friend ReadOnly MouseX As Integer
    Friend ReadOnly MouseY As Integer
    Friend ReadOnly MouseAbsX As Integer
    Friend ReadOnly MouseAbsY As Integer
    Friend ReadOnly MouseAddX As Integer
    Friend ReadOnly MouseAddY As Integer
    Friend ReadOnly IsInMouse As Boolean
    Friend ReadOnly IsInMouseNorth As Boolean
    Friend ReadOnly IsInMouseSouth As Boolean
    Friend ReadOnly IsInMouseEast As Boolean
    Friend ReadOnly IsInMouseWest As Boolean

    Friend IsMouseButtonOn As Boolean
    Friend IsShaked As Boolean = False
    Friend IsSwiped As Boolean = False
    Friend timeStamp As DateTime


    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks>
    ''' コントロール渡しのコンストラクタ
    ''' コントロールの現在の値を取得したインスタンスが生成される。
    ''' </remarks>
    Public Sub New(ByVal control As Windows.Forms.Control)

        Me.X = control.Location.X
        Me.Y = control.Location.Y
        Me.Width = control.Width
        Me.Height = control.Height

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks>
    ''' パラメータ渡しのコンストラクタ
    ''' </remarks>
    Public Sub New(Optional ByVal x As Integer = Nothing, _
                    Optional ByVal y As Integer = Nothing, _
                    Optional ByVal width As Integer = Nothing, _
                    Optional ByVal height As Integer = Nothing)

        Me.X = x
        Me.Y = y
        Me.Width = width
        Me.Height = height

    End Sub


    ''' <summary>
    ''' イベント検知判定用に詳細情報を取得するコンストラクタ
    ''' Xb.Ui.Effect.Monitorで使用する。
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal control As Windows.Forms.Control, _
                    ByVal before As Xb.Forms.ViewParams)

        If (control Is Nothing) Then Return

        Me.timeStamp = Now()
        Me.X = control.Location.X
        Me.Y = control.Location.Y
        Me.Width = control.Width
        Me.Height = control.Height

        'マウスの位置情報を取得
        Dim absPoint, innerPoint As Drawing.Point, _
            widthSlicePercent, heightSlicePercent As Double, _
            widthSlicePixel, heightSlicePixel As Integer

        absPoint = Windows.Forms.Cursor.Position
        Try
            innerPoint = control.PointToClient(absPoint)
        Catch ex As Exception
            Return
        End Try

        Me.MouseAbsX = absPoint.X
        Me.MouseAbsY = absPoint.Y
        Me.MouseX = innerPoint.X
        Me.MouseY = innerPoint.Y


        '上下左右のマウス位置検出範囲を、幅・高さのパーセンテージ設定から取得する。
        widthSlicePercent = 0.2     '左右偏り検出は幅の   20%
        heightSlicePercent = 0.3    '上下偏り検出は高さの 30%
        widthSlicePixel = CInt(control.Width * widthSlicePercent)       '幅偏り率に基づく、幅ピクセル数
        heightSlicePixel = CInt(control.Height * heightSlicePercent)    '高さ偏り率に基づく、高さピクセル数

        Me.IsInMouse = Me.InMouse(innerPoint, 0, 0, control.Width, control.Height)
        Me.IsInMouseNorth = Me.InMouse(innerPoint, 0, 0, control.Width, heightSlicePixel)
        Me.IsInMouseSouth = Me.InMouse(innerPoint, 0, (control.Height - heightSlicePixel), control.Width, control.Height)
        Me.IsInMouseWest = Me.InMouse(innerPoint, 0, 0, widthSlicePixel, control.Height)
        Me.IsInMouseEast = Me.InMouse(innerPoint, (control.Width - widthSlicePixel), 0, control.Width, control.Height)

        'Me.IsMouseButtonOn は、Xb.Ui.Monitorから書き込む。


        If (before isNot Nothing) Then
            Me.AddX = Me.X - before.X
            Me.AddY = Me.Y - before.Y
            Me.MouseAddX = Me.MouseX - before.MouseX
            Me.MouseAddY = Me.MouseY - before.MouseY
        Else
            Me.AddX = 0
            Me.AddY = 0
            Me.MouseAddX = 0
            Me.MouseAddY = 0
        End If

    End Sub


    ''' <summary>
    ''' 渡し値座標が、渡し値範囲に入っているか否かを検証する。
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="startX"></param>
    ''' <param name="startY"></param>
    ''' <param name="endX"></param>
    ''' <param name="endY"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InMouse(ByVal point As Drawing.Point, _
                                ByVal startX As Integer, _
                                ByVal startY As Integer, _
                                ByVal endX As Integer, _
                                ByVal endY As Integer) As Boolean
        Return (startX <= point.X And point.X <= endX _
            And startY <= point.Y And point.Y <= endY)
    End Function


    ''' <summary>
    ''' 未入力パラメータを、渡し値コントロールの値で上書きする。
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Function Format(ByVal control As Windows.Forms.Control) As Xb.Forms.ViewParams

        If (Me.X = Integer.MinValue) Then Me.X = control.Location.X
        If (Me.Y = Integer.MinValue) Then Me.Y = control.Location.Y
        If (Me.Width = Integer.MinValue) Then Me.Width = control.Width
        If (Me.Height = Integer.MinValue) Then Me.Height = control.Height

        Return Me

    End Function

End Class
