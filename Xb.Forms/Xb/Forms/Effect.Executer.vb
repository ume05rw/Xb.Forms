Option Strict On

'Xb.Control.Effectクラスの分割定義
Partial Public Class Effect

    ''' <summary>
    ''' アニメーション効果プロセス管理クラス
    ''' </summary>
    ''' <remarks>
    ''' Form内に表示されているコントロールをアニメーションするクラス。
    ''' Form自体をアニメーションするには AnimateWindowメソッドを使用する。
    ''' </remarks>
    Public Class Executer
        Implements IDisposable

        Public Event Finished As EventHandler

        Private Const Interval As Integer = CInt(1000 / 29) '29fps

        Private ReadOnly _control As Windows.Forms.Control
        Private ReadOnly _form As Windows.Forms.Form
        Private ReadOnly _fromParam As Xb.Forms.ViewParams
        Private ReadOnly _toParam As Xb.Forms.ViewParams
        Private ReadOnly _expectedUpdates As Integer
        Private ReadOnly _timer As System.Threading.Timer
        Private ReadOnly _addLeft As Double
        Private ReadOnly _addTop As Double
        Private ReadOnly _addWidth As Double
        Private ReadOnly _addHeight As Double
        Private _count As Integer
        Private _isOver As Boolean
        Private _dummyControlInUiThread As Windows.Forms.Control


        ''' <summary>
        ''' アニメーション処理が完了したか否か
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' 当初名"IsComplete"、正常終了した名前っぽいのでとりやめ。
        ''' 正常終了した／Disposeされた、のどちらにも関わらずアニメーションが動作していない状況を表す。
        ''' </remarks>
        Public ReadOnly Property IsOver() As Boolean
            Get
                Return Me._isOver
            End Get
        End Property


        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="control"></param>
        ''' <param name="toParam"></param>
        ''' <param name="duration"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal control As Windows.Forms.Control, _
                                ByVal toParam As Xb.Forms.ViewParams, _
                                ByVal duration As Integer)

            Me._dummyControlInUiThread = control
            Me._control = control
            Me._form = Xb.Forms.Util.GetForm(control)

            Me._fromParam = New Xb.Forms.ViewParams(Me._control)
            Me._toParam = toParam.Format((Me._control))

            '何回描画処理を行うかの予定回数。
            Me._expectedUpdates = CInt(duration / Interval)

            'アニメーション１コマ単位の更新量を計算する。
            Me._addLeft = (toParam.X - Me._fromParam.X) / Me._expectedUpdates
            Me._addTop = (toParam.Y - Me._fromParam.Y) / Me._expectedUpdates
            Me._addWidth = (toParam.Width - Me._fromParam.Width) / Me._expectedUpdates
            Me._addHeight = (toParam.Height - Me._fromParam.Height) / Me._expectedUpdates

            'アニメーションを実行する。
            Me._count = 0
            Dim dlg As System.Threading.TimerCallback = New System.Threading.TimerCallback(AddressOf Tick)
            Me._timer = New System.Threading.Timer(dlg, Nothing, 0, Interval)
            Me._isOver = False

        End Sub


        ''' <summary>
        ''' ロールバック専用コンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub New(ByVal control As Windows.Forms.Control, _
                        ByVal form As Windows.Forms.Form, _
                        ByVal fromParam As Xb.Forms.ViewParams, _
                        ByVal toParam As Xb.Forms.ViewParams, _
                        ByVal expectedUpdates As Integer, _
                        ByVal count As Integer)

            Me._dummyControlInUiThread = control
            Me._control = control
            Me._form = form
            Me._fromParam = fromParam
            Me._toParam = toParam
            Me._expectedUpdates = expectedUpdates
            Me._count = count

            'アニメーション１コマ単位量を計算する。
            Me._addLeft = (toParam.X - fromParam.X) / Me._expectedUpdates
            Me._addTop = (toParam.Y - fromParam.Y) / Me._expectedUpdates
            Me._addWidth = (toParam.Width - fromParam.Width) / Me._expectedUpdates
            Me._addHeight = (toParam.Height - fromParam.Height) / Me._expectedUpdates

            'アニメーションを実行する。
            Dim dlg As System.Threading.TimerCallback = New System.Threading.TimerCallback(AddressOf Tick)
            Me._timer = New System.Threading.Timer(dlg, Nothing, 0, Interval)
            Me._isOver = False
        End Sub


        ''' <summary>
        ''' UIスレッドで渡し値イベントをレイズする。
        ''' </summary>
        ''' <param name="eventType"></param>
        ''' <param name="args"></param>
        ''' <remarks></remarks>
        Private Sub RaiseEventOnUiThread(ByVal eventType As Object, ByVal args As Object)
            If (Me._dummyControlInUiThread.InvokeRequired) Then
                Dim dlg As Xb.App.Thread.EventDelegate = New Xb.App.Thread.EventDelegate(AddressOf Me.RaiseEventOnUiThread)
                Try
                    Me._dummyControlInUiThread.Invoke(dlg, New Object() {eventType, args})
                Catch ex As Exception
                End Try
            Else
                Select Case eventType.ToString()
                    Case "Finished"
                        RaiseEvent Finished(Me, New EventArgs())
                    Case Else
                        '何もしない。
                End Select
            End If
        End Sub


        ''' <summary>
        ''' アニメーション１コマ単位の処理
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <remarks></remarks>
        Private Sub Tick(ByVal obj As Object)
            'UIスレッド上で描画する。
            If (Me._form.InvokeRequired) Then
                Dim dlg As System.Threading.TimerCallback = New System.Threading.TimerCallback(AddressOf Tick)
                Try
                    Me._form.Invoke(dlg, New Object() {obj})
                Catch ex As Exception
                    Xb.Util.Out("Xb.Ui.Effect.Executer.Tick: UIスレッドでのInvokeに失敗しました：" & ex.Message)
                End Try

            Else
                If (Me._count < Me._expectedUpdates) Then
                    'アニメーション途中
                    Me._count += 1
                    If (Me._addLeft <> 0) Then Me._control.Left = CInt(Me._fromParam.X + (Me._addLeft * Me._count))
                    If (Me._addTop <> 0) Then Me._control.Top = CInt(Me._fromParam.Y + (Me._addTop * Me._count))
                    If (Me._addWidth <> 0) Then Me._control.Width = CInt(Me._fromParam.Width + (Me._addWidth * Me._count))
                    If (Me._addHeight <> 0) Then Me._control.Height = CInt(Me._fromParam.Height + (Me._addHeight * Me._count))

                    Me._form.Refresh()
                Else
                    'アニメーション完了
                    Me._timer.Dispose()
                    If (Me._addLeft <> 0) Then Me._control.Left = Me._toParam.X
                    If (Me._addTop <> 0) Then Me._control.Top = Me._toParam.Y
                    If (Me._addWidth <> 0) Then Me._control.Width = Me._toParam.Width
                    If (Me._addHeight <> 0) Then Me._control.Height = Me._toParam.Height
                    Me._form.Refresh()
                    Me._isOver = True

                    '処理終了イベントをレイズする。
                    Me.RaiseEventOnUiThread("Finished", Nothing)
                    Me.Dispose()
                End If
            End If

        End Sub


        ''' <summary>
        ''' 元の表示に戻す。
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Rollback() As Xb.Forms.Effect.Executer
            '現在のアニメーションを停止
            Me._timer.Dispose()

            '起動時の渡し値と現在のカウントで、逆方向のアニメーションオブジェクトを生成する。
            Return New Xb.Forms.Effect.Executer(Me._control, _
                                                 Me._form, _
                                                 Me._toParam, _
                                                 Me._fromParam, _
                                                 Me._expectedUpdates, _
                                                 (Me._expectedUpdates - Me._count))
        End Function


#Region "IDisposable Support"
        Private _disposedValue As Boolean ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me._disposedValue Then
                If disposing Then
                    If (Not Me._timer Is Nothing) Then
                        Me._timer.Dispose()
                    End If
                    Me._isOver = True
                End If
            End If
            Me._disposedValue = True
        End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Class
