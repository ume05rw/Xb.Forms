Option Strict On

''' <summary>
''' UIエフェクト管理クラス
''' </summary>
''' <remarks></remarks>
Partial Public Class Effect


    <Flags()> _
    Public Enum WindowFlags
        AwHorPositive = &H1
        AwHorNegative = &H2
        AwVerPositive = &H4
        AwVerNegative = &H8
        AwCenter = &H10
        AwHide = &H10000
        AwActivate = &H20000
        AwSlide = &H40000
        AwBlend = &H80000
    End Enum

    <System.Runtime.InteropServices.DllImport("user32.dll")> _
    Private Shared Function AnimateWindow(ByVal windowHandle As IntPtr, _
                                            ByVal time As Integer, _
                                            ByVal animateFlags As Xb.Forms.Effect.WindowFlags) As Boolean
    End Function

    'ウインドウのフェードイン
    'Dim form As New Form
    'Xb.Ui.Effect.AnimateWindow(form.Handle, 200, AnimateWindowFlags.AW_ACTIVATE Or AnimateWindowFlags.AW_BLEND)
    'form.Show()

    'アラートウインドウのエフェクト例
    'AnimateWindow(f.Handle, 200, AnimateWindowFlags.AW_SLIDE Or AnimateWindowFlags.AW_VER_NEGATIVE)

    Public Shared Sub FadeinWindow(ByVal form As Windows.Forms.Form, _
                                    Optional ByVal duration As Integer = 300)

        If (form Is Nothing) Then Throw New ArgumentException("Formが検知できません。")

        form.Hide()
        Xb.Forms.Effect.AnimateWindow(form.Handle, duration, WindowFlags.AwActivate Or WindowFlags.AwBlend)
        form.Show()
    End Sub

    Public Shared Sub SlideUpWindow(ByVal form As Windows.Forms.Form, _
                                    Optional ByVal duration As Integer = 300)
        If (form Is Nothing) Then Throw New ArgumentException("Formが検知できません。")

        form.Hide()
        Xb.Forms.Effect.AnimateWindow(form.Handle, duration, WindowFlags.AwSlide Or WindowFlags.AwVerNegative)
        form.Show()
    End Sub
End Class
