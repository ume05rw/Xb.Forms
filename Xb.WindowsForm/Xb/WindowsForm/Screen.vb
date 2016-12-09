Option Strict On

''' <summary>
''' ディスプレイ関連関数群
''' </summary>
''' <remarks></remarks>
Public Partial Class Screen

            ''' <summary>
    ''' 渡し値フォームが表示されているスクリーン番号を取得する。
    ''' </summary>
    ''' <param name="form"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetActiveScreenNumber(ByRef form As Windows.Forms.Form) As Integer

        Dim screens() As Windows.Forms.Screen, _
            formX, formY, i As Integer

        screens = Windows.Forms.Screen.AllScreens()
        formX = form.Location.X
        formY = form.Location.Y

        For i = 0 To screens.Length - 1
            If ( _
                (screens(i).Bounds.X <= formX) _
                And (screens(i).Bounds.X + screens(i).Bounds.Width > formX) _
                And (screens(i).Bounds.Y <= formY) _
                And (screens(i).Bounds.Y + screens(i).Bounds.Height > formY) _
            ) Then
                Return i
            End If
        Next

        Xb.Util.Out("Xb.Device.Screen.GetActiveScreenNumber: アクティブなスクリーンが見つかりません。")
        Throw New ApplicationException("アクティブなスクリーンが見つかりません。")

    End Function


    ''' <summary>
    ''' 渡し値フォームが表示されているスクリーンオブジェクトを取得する。
    ''' </summary>
    ''' <param name="form"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetActiveScreen(ByRef form As Windows.Forms.Form) As Windows.Forms.Screen

        Dim screens() As Windows.Forms.Screen

        screens = Windows.Forms.Screen.AllScreens()

        Return screens(GetActiveScreenNumber(form))

    End Function


    ''' <summary>
    ''' 全画面のスクリーンキャプチャ画像を取得する。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' プライマリモニタでなく、仮想画面全面をキャプチャする。
    ''' </remarks>
    Public Shared Function Capture() As Drawing.Image

        Dim targetBounds As Drawing.Rectangle, _
            image As Drawing.Image

        targetBounds = Windows.Forms.SystemInformation.VirtualScreen
        image = New Drawing.Bitmap(targetBounds.Width, targetBounds.Height)

        Using graph As Drawing.Graphics = Drawing.Graphics.FromImage(image)
            graph.CopyFromScreen(targetBounds.Location, New Drawing.Point(0, 0), image.Size)
        End Using

        Return image

    End Function


    ''' <summary>
    ''' 渡し値画面のキャプチャ画像を取得する。
    ''' </summary>
    ''' <param name="form"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Capture(ByVal form As System.Windows.Forms.Form) As Drawing.Image

        If (form Is Nothing) Then
            Xb.Util.Out("Xb.Device.Screen.Capture: フォームが検出出来ません。")
            Throw New ArgumentException("フォームが検出出来ません。")
        End If

        Dim image As Drawing.Bitmap = New Drawing.Bitmap(form.Width, form.Height)
        form.DrawToBitmap(image, New Drawing.Rectangle(0, 0, form.Width, form.Height))

        Return image

    End Function

End Class



