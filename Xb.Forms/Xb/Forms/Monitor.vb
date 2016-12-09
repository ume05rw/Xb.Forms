Option Strict On

Imports System.Runtime.InteropServices


''' <summary>
''' 指定コントロールの状態監視・イベントホストクラス
''' </summary>
''' <remarks>
''' グローバルフックのメッセージを元に、マウス／指定コントロールの状態を監視するクラス。
''' コントロールのネストや他アプリケーションによるコントロール支配など、
''' 通常のイベントが発生しない場合のイベントレイズ代替手段。
''' 
''' マウスイベント
'''   MouseEnter
'''   MouseLeave
'''
''' カスタムイベント
'''   MouseEnterNorth: マウスが対象の 上部 30% の範囲に入った
'''   MouseLeaveNorth: マウスが対象の 上部 30% の範囲から出た
'''   MouseEnterSouth: マウスが対象の 下部 30% の範囲に入った
'''   MouseLeaveSouth: マウスが対象の 下部 30% の範囲から出た
'''   MouseEnterEast:  マウスが対象の 右部 20% の範囲に入った
'''   MouseLeaveEast:  マウスが対象の 右部 20% の範囲から出た
'''   MouseEnterWest:  マウスが対象の 左部 20% の範囲に入った
'''   MouseLeaveWest:  マウスが対象の 左部 20% の範囲から出た
''' 
'''   Shake: コントロールの座標が短時間に前後左右に動いた
''' 
''' TODO:実装したいイベント
'''   MouseClick
'''   MouseDoubleClick
'''   Drag:  マウスがコントロール上で左ボタンON状態に変化し、ボタンを維持したままマウスが動いた
'''   Swipe: マウスがコントロール上で左ボタンON状態に変化し、ボタンを維持したままコントロール上を一定の方向に進んだ
''' 
'''   ScreenChanged: 表示するスクリーンが変わった。
''' </remarks>
Public Class Monitor
    Implements IDisposable


    ''' <summary>
    ''' アプリケーション定義のフックプロシージャ"lpfn"をフックチェーン内にインストールします。
    ''' </summary>
    ''' <param name="idHook">フックタイプ</param>
    ''' <param name="lpfn">フックプロシージャ</param>
    ''' <param name="hInstance">アプリケーションインスタンスのハンドル</param>
    ''' <param name="threadId">スレッドの識別子</param>
    ''' <returns>フックプロシージャのハンドル</returns>
    ''' <remarks>
    ''' 
    ''' SetWindowsHookEx
    ''' http://msdn.microsoft.com/ja-jp/library/cc430103.aspx
    ''' 
    ''' マウス対象のグローバルフックなので、idHook = WH_MOUSE_LL(=14) になる。
    ''' 
    ''' フック時のコールバック関数が LowLevelMouseProc(=HookProcedureDelegate)。
    ''' http://msdn.microsoft.com/ja-jp/library/cc430021.aspx
    ''' 
    ''' </remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function SetWindowsHookEx(ByVal idHook As Integer, _
                                        ByVal lpfn As HookProcedureDelegate, _
                                        ByVal hInstance As IntPtr, _
                                        ByVal threadId As Integer) As IntPtr
    End Function


    ''' <summary>
    ''' SetWindowsHookEx 関数を使ってフックチェーン内にインストールされたフックプロシージャを削除します。
    ''' </summary>
    ''' <param name="idHook">削除対象のフックプロシージャのハンドル</param>
    ''' <returns>関数が成功すると 0 以外の値、失敗すると 0 が返ります。</returns>
    ''' <remarks>
    ''' UnhookWindowsHookEx
    ''' http://msdn.microsoft.com/ja-jp/library/cc430120.aspx
    ''' </remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function UnhookWindowsHookEx(ByVal idHook As IntPtr) As Boolean
    End Function


    ''' <summary>
    ''' 現在のフックチェーン内の次のフックプロシージャに、フック情報を渡します。
    ''' </summary>
    ''' <param name="idHook">現在のフックのハンドル</param>
    ''' <param name="nCode">フックプロシージャに渡すフックコード</param>
    ''' <param name="wParam">フックプロシージャに渡す値</param>
    ''' <param name="lParam">フックプロシージャに渡す値</param>
    ''' <returns>フックチェーン内の次のフックプロシージャの戻り値</returns>
    ''' <remarks>
    ''' CallNextHookEx
    ''' http://msdn.microsoft.com/ja-jp/library/cc429591.aspx
    ''' </remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, _
    CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function CallNextHookEx(ByVal idHook As IntPtr, _
                                        ByVal nCode As Integer, _
                                        ByVal wParam As IntPtr, _
                                        ByVal lParam As IntPtr) As IntPtr
    End Function


    ''' <summary>
    ''' 呼び出し側プロセスのアドレス空間に該当ファイルがマップされている場合、
    ''' 指定されたモジュール名のモジュールハンドルを返します。
    ''' </summary>
    ''' <param name="lpModuleName">モジュール名</param>
    ''' <returns>関数が成功時はモジュールのハンドル、失敗時はNullが返る。</returns>
    ''' <remarks>
    ''' GetModuleHandle
    ''' http://msdn.microsoft.com/ja-jp/library/cc429129.aspx
    ''' </remarks>
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function


    ''' <summary>
    ''' LowLevelMouseProc関数を示すデリゲート
    ''' </summary>
    ''' <param name="nCode">フックコード</param>
    ''' <param name="wParam">
    ''' メッセージ識別子
    '''   WM_MOUSEMOVE       : 0x0200 (512) : マウス移動
    '''   WM_LBUTTONDOWN     : 0x0201 (513) : マウス左ボタンを押した
    '''   WM_LBUTTONUP       : 0x0202 (514) : マウス左ボタンを離した
    '''   WM_LBUTTONDBLCLK   : 0x0203 (515) : マウス左ボタンをダブルクリック
    '''   WM_RBUTTONDOWN     : 0x0204 (516) : マウス右ボタンを押した
    '''   WM_RBUTTONUP       : 0x0205 (517) : マウス右ボタンを離した
    '''   WM_RBUTTONDBLCLK   : 0x0206 (518) : マウス右ボタンをダブルクリック
    '''   WM_MOUSEWHEEL      : 0x020A (522) : マウスホイールが回転した。(上下どちらかは不明。)
    ''' 　
    '''  　 ↑※左右ダブルクリックを検出するには、対象となる画面が必要。
    ''' 　　↑※グローバルフックには画面が存在しないため、ダブルクリック発生しないようだ。
    ''' </param>
    ''' <param name="lParam">メッセージデータ(=MSLLHOOKSTRUCT構造体)</param>
    ''' <returns>(強く推奨)フックチェーン内の次のフックプロシージャへの戻り値</returns>
    ''' <remarks>
    ''' LowLevelMouseProc
    ''' http://msdn.microsoft.com/ja-jp/library/cc430021.aspx
    ''' </remarks>
    Public Delegate Function HookProcedureDelegate(nCode As Integer, _
                                                    wParam As IntPtr, _
                                                    lParam As IntPtr) As IntPtr


    ''' <summary>
    ''' 低レベルマウス入力イベント情報構造体
    ''' </summary>
    ''' <remarks>
    ''' MSLLHOOKSTRUCT structure
    ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms644970%28v=vs.85%29.aspx
    ''' </remarks>
    <StructLayout(LayoutKind.Sequential)> _
    Public Class MsllHookStruct

        ''' <summary>
        ''' スクリーン上のX, Y座標
        ''' </summary>
        ''' <remarks></remarks>
        Public point As PointStruct

        ''' <summary>
        ''' マウス情報値
        ''' </summary>
        ''' <remarks>
        ''' WM_MOUSEWHEEL      : 0x020A (522) : ホイールが回転した
        ''' WM_XBUTTONDOWN     : 0x020B (523) : Xボタンを押した
        ''' WM_XBUTTONUP       : 0x020C (524) : Xボタンを離した
        ''' WM_XBUTTONDBLCLK   : 0x020D (525) : Xボタンをダブルクリック
        ''' WM_NCXBUTTONDOWN   : 0x00AB (171) : Xボタンを押した
        ''' WM_NCXBUTTONUP     : 0x00AC (172) : Xボタンを離した
        ''' WM_NCXBUTTONDBLCLK : 0x00AD (173) : Xボタンをダブルクリック
        ''' </remarks>
        Public mouseData As UInteger

        ''' <summary>
        ''' イベント割り込みフラグ
        ''' </summary>
        ''' <remarks></remarks>
        Public flags As UInteger

        ''' <summary>
        ''' タイムスタンプ
        ''' </summary>
        ''' <remarks></remarks>
        Public time As UInteger

        ''' <summary>
        ''' その他追加メッセージ
        ''' </summary>
        ''' <remarks></remarks>
        Public dwExtraInfo As IntPtr
    End Class


    ''' <summary>
    ''' X, Y座標を示す構造体
    ''' </summary>
    ''' <remarks>
    ''' POINT structure
    ''' http://msdn.microsoft.com/en-us/library/windows/desktop/dd162805%28v=vs.85%29.aspx
    ''' </remarks>
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure PointStruct
        Public x As Integer
        Public y As Integer
    End Structure


    'フックするメッセージタイプ、低レベルマウスイベントを示す。SetWindowsHookExで定義される。
    'http://msdn.microsoft.com/ja-jp/library/cc430103.aspx
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202


    Private Const HistoryCount As Integer = 10
    Private Const ClickDelay As Integer = 400       'クリック検出から X MSec 待ってイベント判定する。
    Private Const ClickThreshold As Integer = 700   '判定時点の時刻から X MSec 以上古い履歴は判定対象にしない。

    Public Class MouseEventArgs
        Inherits EventArgs

        Public ReadOnly MouseX As Integer
        Public ReadOnly MouseY As Integer
        Public ReadOnly MouseAbsX As Integer
        Public ReadOnly MouseAbsY As Integer
        Public ReadOnly InNorth As Boolean
        Public ReadOnly InSouth As Boolean
        Public ReadOnly InEast As Boolean
        Public ReadOnly InWest As Boolean

        Public Sub New(ByVal mouseX As Integer, _
                        ByVal mouseY As Integer, _
                        ByVal mouseAbsX As Integer, _
                        ByVal mouseAbsY As Integer, _
                        ByVal inNorth As Boolean, _
                        ByVal inSouth As Boolean, _
                        ByVal inEast As Boolean, _
                        ByVal inWest As Boolean)
            Me.MouseX = mouseX
            Me.MouseY = mouseY
            Me.MouseAbsX = mouseAbsX
            Me.MouseAbsY = mouseAbsY
            Me.InNorth = inNorth
            Me.InSouth = inSouth
            Me.InEast = inEast
            Me.InWest = inWest
        End Sub
    End Class

    ''' <summary>
    ''' マウスドラッグイベントの引数クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MouseDragEventArgs
        Inherits EventArgs

        Public ReadOnly AddX As Integer
        Public ReadOnly AddY As Integer
        Public ReadOnly MouseX As Integer
        Public ReadOnly MouseY As Integer
        Public ReadOnly MouseAbsX As Integer
        Public ReadOnly MouseAbsY As Integer
        Public ReadOnly InNorth As Boolean
        Public ReadOnly InSouth As Boolean
        Public ReadOnly InEast As Boolean
        Public ReadOnly InWest As Boolean

        Public Sub New(ByVal addX As Integer, _
                        ByVal addY As Integer, _
                        ByVal mouseX As Integer, _
                        ByVal mouseY As Integer, _
                        ByVal mouseAbsX As Integer, _
                        ByVal mouseAbsY As Integer, _
                        ByVal inNorth As Boolean, _
                        ByVal inSouth As Boolean, _
                        ByVal inEast As Boolean, _
                        ByVal inWest As Boolean)
            Me.AddX = addX
            Me.AddY = addY
            Me.MouseX = mouseX
            Me.MouseY = mouseY
            Me.MouseAbsX = mouseAbsX
            Me.MouseAbsY = mouseAbsY
            Me.InNorth = inNorth
            Me.InSouth = inSouth
            Me.InEast = inEast
            Me.InWest = inWest
        End Sub
    End Class

    'イベント用デリゲート
    Public Delegate Sub MoveEventHandler(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Delegate Sub DragEventHandler(ByVal sender As Object, ByVal e As MouseDragEventArgs)
    Public Delegate Sub ClickEventHandler(ByVal sender As Object, ByVal e As MouseEventArgs)
    Public Delegate Sub DoubleClickEventHandler(ByVal sender As Object, ByVal e As MouseEventArgs)

    Public Event MouseClicked As ClickEventHandler
    Public Event MouseDoubleClicked As DoubleClickEventHandler
    Public Event MouseMove As MoveEventHandler
    Public Event MouseDragged As DragEventHandler
    Public Event MouseEnter As EventHandler
    Public Event MouseLeave As EventHandler
    Public Event MouseEnterNorth As EventHandler
    Public Event MouseLeaveNorth As EventHandler
    Public Event MouseEnterSouth As EventHandler
    Public Event MouseLeaveSouth As EventHandler
    Public Event MouseEnterEast As EventHandler
    Public Event MouseLeaveEast As EventHandler
    Public Event MouseEnterWest As EventHandler
    Public Event MouseLeaveWest As EventHandler
    Public Event Shaked As EventHandler
    'Public Event Swiped As EventHandler

    Private _history As List(Of Xb.Forms.ViewParams)
    Private _control As Windows.Forms.Control
    Private _form As Windows.Forms.Form
    Private _mouseHandle As IntPtr = IntPtr.Zero
    Private _delegate As Xb.Forms.Monitor.HookProcedureDelegate

    Private _isMouseButtonOn As Boolean
    Private _mouseButtonStartPoint As Xb.Forms.ViewParams

    Private _timer As Windows.Forms.Timer
    Private _lastClickEventTime As DateTime = DateTime.MinValue

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal control As Windows.Forms.Control)

        '入力チェック
        If (control Is Nothing) Then
            Xb.Util.Out("Xb.Ui.Monitor.New: 渡し値コントロールが検出出来ません。")
            Throw New ArgumentException("渡し値コントロールが検出出来ません。")
        End If

        Me._history = New List(Of Xb.Forms.ViewParams)()
        Me._control = control
        Me._form = Xb.Forms.Util.GetForm(Me._control)
        Me._delegate = New HookProcedureDelegate(AddressOf OnMouseChanged)
        Me._isMouseButtonOn = False
        Me._mouseButtonStartPoint = Nothing

        'マウスのイベント取得を開始する。
        Using processes As Process = Process.GetCurrentProcess()
            Using mainModule As ProcessModule = processes.MainModule
                Me._mouseHandle = SetWindowsHookEx(Monitor.WH_MOUSE_LL, _
                                                Me._delegate, _
                                                GetModuleHandle(mainModule.ModuleName), 0)
            End Using
        End Using

        If (Me._mouseHandle = IntPtr.Zero) Then
            Xb.Util.Out("Xb.Ui.Monitor.New: グローバルフック取得に失敗しました。")
            Throw New ApplicationException("グローバルフック取得に失敗しました。")
        End If

        'クリックイベントの遅延判定処理用に生成して待機させておく。
        Me._timer = New Windows.Forms.Timer()
        Me._timer.Interval = ClickDelay
        AddHandler Me._timer.Tick, AddressOf Me.CheckClickEvent

    End Sub


    ''' <summary>
    ''' コントロールとマウスの座標履歴を取り、イベント判定を行う。
    ''' </summary>
    ''' <param name="nCode"></param>
    ''' <param name="wParam"></param>
    ''' <param name="lParam"></param>
    ''' <remarks>
    ''' マウスを動かす都度実行される。但しマウス操作が無い限り実行されない。
    ''' </remarks>
    Private Function OnMouseChanged(ByVal nCode As Integer, _
                                    ByVal wParam As IntPtr, _
                                    ByVal lParam As IntPtr) As IntPtr

        Dim messageId As Integer = wParam.ToInt32()

        'マウスのボタン押下状態を保持しておく。
        If (messageId = Monitor.WM_LBUTTONDOWN) Then
            Me._isMouseButtonOn = True
            Xb.Util.Out("--- MouseButtonOn ---")
        ElseIf (messageId = Monitor.WM_LBUTTONUP) Then
            Me._isMouseButtonOn = False
            Me._mouseButtonStartPoint = Nothing
            Xb.Util.Out("--- MouseButtonOff ---")
        End If

        'TODO: フォームがアクティブでないとき、いきなりドラッグが出来ない。
        'TODO: マウスとフォームの間に他のフォームが無いことを判定出来ればいいが。
        '監視フォームがアクティブのときのみイベントを発生させる。
        If (Me._form.Equals(Windows.Forms.Form.ActiveForm)) Then

            If (nCode >= 0) Then
                '履歴情報を更新する。
                Me.UpdateHistory(messageId)

                'イベント検証処理
                If (Me._history.Count > 2) Then Me.CheckEvent(messageId)
            End If
        End If

        'フックチェーン内の次のフックプロシージャへ現在のフックを送る。
        '下記を書かないでおくと、この処理以降のマウスフック待機中プロシージャがフックを受け取れない。
        Return CallNextHookEx(Me._mouseHandle, nCode, wParam, lParam)

    End Function


    ''' <summary>
    ''' マウスイベントの履歴情報を更新する。
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <remarks></remarks>
    Private Sub UpdateHistory(ByVal messageId As Integer)

        '現在のコントロールの座標・表示領域情報を履歴に追加する。
        Dim current As Xb.Forms.ViewParams

        If (Me._history.Count <= 0) Then
            current = New Xb.Forms.ViewParams(Me._control, Nothing)
        Else
            current = New Xb.Forms.ViewParams(Me._control, Me._history.Item(Me._history.Count - 1))
        End If


        If (Me._isMouseButtonOn And Me._mouseButtonStartPoint is Nothing) Then
            'マウスボタンがONのとき、かつマウスON時点の座標が無いとき、保持しておく。
            'Xb.App.Out("--- Detect Button On ---")
            Me._mouseButtonStartPoint = current
        ElseIf (Not Me._isMouseButtonOn And Me._mouseButtonStartPoint isnot Nothing) Then
            'マウスボタンがOFFのとき、マウス座標オブジェクトを空にしておく
            'Xb.App.Out("--- Detect Button Release ---")
            Me._mouseButtonStartPoint = Nothing
        End If

        current.IsMouseButtonOn = Me._isMouseButtonOn
        Me._history.Add(current)

        '履歴は直近10件のみ保持するように。
        If (Me._history.Count > HistoryCount) Then Me._history.RemoveRange(0, Me._history.Count - HistoryCount)

    End Sub


    ''' <summary>
    ''' 現在までの操作イベント履歴から、カスタムイベントを発生させる。
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <remarks></remarks>
    Private Sub CheckEvent(ByVal messageId As Integer)

        Dim current, before As Xb.Forms.ViewParams, _
            x, y, addX, addY, absAddX, absAddY, mouseX, mouseY, _
            mouseAddX, mouseAddY, mouseAbsAddX, mouseAbsAddY As Double, _
            isKeepInMouse, isShaked, isSwiped As Boolean

        current = Me._history.Item(Me._history.Count - 1)
        before = Me._history.Item(Me._history.Count - 2)

        '直近10回分の平均値。一応初期化。
        'コントロールの座標・運動量
        x = 0           'X座標の平均値
        y = 0           'Y座標の平均値
        addX = 0        'X軸増分の平均値
        addY = 0        'Y軸増分の平均値
        absAddX = 0     'X軸増分絶対値の平均値
        absAddY = 0     'Y軸増分絶対値の平均値

        'マウスの座標・運動量
        mouseX = 0          'コントロールと同じ並び順
        mouseY = 0          '
        mouseAddX = 0       '
        mouseAddY = 0       '
        mouseAbsAddX = 0    '
        mouseAbsAddY = 0    '

        '直近10回の中でShaked, Swipedイベントが発生しているか否か
        isShaked = False
        isSwiped = False

        'マウスがコントロール範囲に存在し続けたか否かのフラグ
        isKeepInMouse = True '初期値。履歴に一つでも範囲外があると、False

        '処理中に別スレッドで履歴に要素追加されてもいいように。
        'されるのかしらんけど。
        For i As Integer = 0 To Me._history.Count - 1
            x += Me._history.Item(i).X
            y += Me._history.Item(i).Y
            addX += Me._history.Item(i).AddX
            addY += Me._history.Item(i).AddY
            absAddX += Math.Abs(Me._history.Item(i).AddX)
            absAddY += Math.Abs(Me._history.Item(i).AddY)

            mouseX += Me._history.Item(i).MouseX
            mouseY += Me._history.Item(i).MouseY
            mouseAddX += Me._history.Item(i).MouseAddX
            mouseAddY += Me._history.Item(i).MouseAddY
            mouseAbsAddX += Math.Abs(Me._history.Item(i).MouseAddX)
            mouseAbsAddY += Math.Abs(Me._history.Item(i).MouseAddY)

            If (Me._history.Item(i).IsShaked) Then isShaked = True
            If (Me._history.Item(i).IsSwiped) Then isSwiped = True

            If (Not Me._history.Item(i).IsInMouse) Then isKeepInMouse = False
        Next

        x = x / Me._history.Count
        y = y / Me._history.Count
        addX = addX / Me._history.Count
        addY = addY / Me._history.Count
        absAddX = absAddX / Me._history.Count
        absAddY = absAddY / Me._history.Count

        '以下、判定結果
        If (messageId = Monitor.WM_MOUSEMOVE) Then
            RaiseEvent MouseMove(Me, New MouseEventArgs(current.MouseX, _
                                                        current.MouseY, _
                                                        current.MouseAbsX, _
                                                        current.MouseAbsY, _
                                                        current.IsInMouseNorth, _
                                                        current.IsInMouseSouth, _
                                                        current.IsInMouseEast, _
                                                        current.IsInMouseWest))
        End If

        If (current.IsInMouse _
            And (messageId = Monitor.WM_LBUTTONUP Or messageId = Monitor.WM_LBUTTONDOWN) _
        ) Then
            'Xb.App.Out("Monitor.CheckEvent_Before TimerValidate.")
            'クリックイベント判定。操作直後でなく、少し待ってから操作履歴を調べ、
            'シングル or ダブルクリックを判定している。
            If (Not Me._timer.Enabled) Then
                'Xb.App.Out("Monitor.CheckEvent_MouseClickEvent-Validate Start.")
                Me._timer.Start()
            End If
        End If
        'AndAlso Me._isMouseButtonOn _
        If (current.IsInMouse _
            AndAlso Me._mouseButtonStartPoint isNot Nothing _
            AndAlso messageId = Monitor.WM_MOUSEMOVE) Then

            'Dim log As String = ""
            'log &= vbCrLf & " マウス起点.MouseAbsX : " & _mouseButtonStartPoint.MouseAbsX
            'log &= vbCrLf & " マウス現点.MouseAbsX : " & current.MouseAbsX
            'log &= vbCrLf & " マウス増分.X         : " & (current.MouseAbsX - Me._mouseButtonStartPoint.MouseAbsX)
            'log &= vbCrLf
            'log &= vbCrLf & " 画面起点.X           : " & _mouseButtonStartPoint.X
            'log &= vbCrLf & " 画面現点.X           : " & current.X
            'log &= vbCrLf & " 画面増分.X           : " & (current.X - Me._mouseButtonStartPoint.X)
            'log &= vbCrLf
            'log &= vbCrLf & " 合計増分.X           : " & (current.MouseAbsX - Me._mouseButtonStartPoint.MouseAbsX) _
            '                                            - (current.X - Me._mouseButtonStartPoint.X)
            'log &= vbCrLf
            'log &= vbCrLf & " マウス起点.MouseAbsY : " & _mouseButtonStartPoint.MouseAbsY
            'log &= vbCrLf & " マウス現点.MouseAbsY : " & current.MouseAbsY
            'log &= vbCrLf & " マウス増分.Y         : " & (current.MouseAbsY - Me._mouseButtonStartPoint.MouseAbsY)
            'log &= vbCrLf
            'log &= vbCrLf & " 画面起点.Y           : " & _mouseButtonStartPoint.Y
            'log &= vbCrLf & " 画面現点.Y           : " & current.Y
            'log &= vbCrLf & " 画面増分.Y           : " & (current.Y - Me._mouseButtonStartPoint.Y)
            'log &= vbCrLf
            'log &= vbCrLf & " 合計増分.Y           : " & (current.MouseAbsY - Me._mouseButtonStartPoint.MouseAbsY) _
            '                                            - (current.Y - Me._mouseButtonStartPoint.Y)
            'log &= vbCrLf
            'Xb.App.Out(log)

            'Xb.App.Out("MouseDragged")
            RaiseEvent MouseDragged(Me, New MouseDragEventArgs( _
                ( _
                    (current.MouseAbsX - Me._mouseButtonStartPoint.MouseAbsX) _
                    - (current.X - Me._mouseButtonStartPoint.X) _
                        ), _
                    ( _
                        (current.MouseAbsY - Me._mouseButtonStartPoint.MouseAbsY) _
                        - (current.Y - Me._mouseButtonStartPoint.Y) _
                    ), _
                    current.MouseX, _
                    current.MouseY, _
                    current.MouseAbsX, _
                    current.MouseAbsY, _
                    current.IsInMouseNorth, _
                    current.IsInMouseSouth, _
                    current.IsInMouseEast, _
                    current.IsInMouseWest _
                ) _
            )
        End If
        If (current.IsInMouse And Not before.IsInMouse) Then
            'Xb.App.Out("MouseEnter")
            RaiseEvent MouseEnter(Me, New EventArgs())
        End If
        If (Not current.IsInMouse And before.IsInMouse) Then
            'Xb.App.Out("MouseLeave")
            RaiseEvent MouseLeave(Me, New EventArgs())
        End If

        If (current.IsInMouseNorth And Not before.IsInMouseNorth) Then
            'Xb.App.Out("MouseEnterNorth")
            RaiseEvent MouseEnterNorth(Me, New EventArgs())
        End If
        If (Not current.IsInMouseNorth And before.IsInMouseNorth) Then
            'Xb.App.Out("MouseLeaveNorth")
            RaiseEvent MouseLeaveNorth(Me, New EventArgs())
        End If

        If (current.IsInMouseSouth And Not before.IsInMouseSouth) Then
            'Xb.App.Out("MouseEnterSouth")
            RaiseEvent MouseEnterSouth(Me, New EventArgs())
        End If
        If (Not current.IsInMouseSouth And before.IsInMouseSouth) Then
            'Xb.App.Out("MouseLeaveSouth")
            RaiseEvent MouseLeaveSouth(Me, New EventArgs())
        End If

        If (current.IsInMouseEast And Not before.IsInMouseEast) Then
            'Xb.App.Out("MouseEnterEast")
            RaiseEvent MouseEnterEast(Me, New EventArgs())
        End If
        If (Not current.IsInMouseEast And before.IsInMouseEast) Then
            'Xb.App.Out("MouseLeaveEast")
            RaiseEvent MouseLeaveEast(Me, New EventArgs())
        End If

        If (current.IsInMouseWest And Not before.IsInMouseWest) Then
            'Xb.App.Out("MouseEnterWest")
            RaiseEvent MouseEnterWest(Me, New EventArgs())
        End If
        If (Not current.IsInMouseWest And before.IsInMouseWest) Then
            'Xb.App.Out("MouseLeaveWest")
            RaiseEvent MouseLeaveWest(Me, New EventArgs())
        End If

        'Shake判定対象：コントロールの移動量
        '１．増分絶対値の平均が大きいこと -> 運動量の総量が大きい
        '２．合計増量が少ないこと -> 当初位置からさほど動いていない
        '  => 振れ幅の多い反復運動と見做す。
        If ( _
            (Not isShaked) _
            And ( _
                (absAddX > 10 And Math.Abs(addX) < 2.2) _
                Or (absAddY > 10 And Math.Abs(addY) < 2.2) _
            ) _
        ) Then
            current.IsShaked = True
            'Xb.App.Out("Shake")
            RaiseEvent Shaked(Me, New EventArgs())
        End If

        '
        '' Swipeは判定が難しいので実装棚上げ。
        ''
        ''Swipe判定対象：マウスの移動量
        ''１．合計増量が大きいこと -> 運動が存在している
        ''２．増分平均の絶対値と増分絶対値の平均が近しいこと -> 運動方向が変わっていない
        ''３．上下方向、左右方向のどちらかに偏った運動であること
        '' => 一定方向への運動と見做す。
        'If ( _
        '    (Not isSwiped) _
        '    And (isKeepInMouse) _
        '    And (Me._isMouseButtonOn) _
        '    And ((Math.Abs(mouseAddX) + Math.Abs(mouseAddY)) > 20) _
        '    And (Math.Abs(mouseAbsAddX - Math.Abs(mouseAddX)) < 2) _
        '    And (Math.Abs(mouseAbsAddY - Math.Abs(mouseAddY)) < 2) _
        '    And (Math.Abs(mouseAddX) < 2 Or Math.Abs(mouseAddY) < 2)
        ') Then
        '    Xb.App.Out("Swipe")
        '    current.IsSwiped = True
        'End If

    End Sub


    ''' <summary>
    ''' クリックイベントの遅延処理。シングル or ダブルクリックの確定を行う。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CheckClickEvent(ByVal sender As Object, ByVal e As EventArgs)
        Me._timer.Stop()
        'Xb.App.Out("Monitor.CheckClickEvent_1")

        If (Me._lastClickEventTime.AddMilliseconds(ClickDelay) > Now()) Then
            '前回クリックイベントから500ミリ秒経過していないとき、イベントを抑制する。
            'Xb.App.Out("CheckClickEvent_2")
            Return
        End If

        Dim changeCount As Integer = 0, _
            buttonOn As Boolean = False, _
            point, lastButtonUp As Xb.Forms.ViewParams

        lastButtonUp = Nothing

        For i As Integer = Me._history.Count - 1 To 0 Step -1
            point = Me._history.Item(i)

            '履歴時刻が 500msecより古い場合は参照しない。
            If (point.timeStamp.AddMilliseconds(ClickThreshold) < Now()) Then
                Continue For
            End If

            'Xb.App.Out("Monitor.CheckClickEvent_2.5 point.IsMouseButtonOn = " & point.IsMouseButtonOn.ToString())
            If (buttonOn <> point.IsMouseButtonOn) Then
                changeCount += 1
                buttonOn = point.IsMouseButtonOn
                lastButtonUp = point
            End If
        Next
        'Xb.App.Out("Monitor.CheckClickEvent_3: Count: " & changeCount)

        Me._lastClickEventTime = Now()
        If (changeCount > 2) Then
            'Xb.App.Out("Monitor.CheckClickEvent_DoubleClick")
            RaiseEvent MouseDoubleClicked(Me, New MouseEventArgs(lastButtonUp.MouseX, _
                                                                lastButtonUp.MouseY, _
                                                                lastButtonUp.MouseAbsX, _
                                                                lastButtonUp.MouseAbsY, _
                                                                lastButtonUp.IsInMouseNorth, _
                                                                lastButtonUp.IsInMouseSouth, _
                                                                lastButtonUp.IsInMouseEast, _
                                                                lastButtonUp.IsInMouseWest))
        ElseIf (changeCount > 0) Then
            'Xb.App.Out("Monitor.CheckClickEvent_SingleClick")
            RaiseEvent MouseClicked(Me, New MouseEventArgs(lastButtonUp.MouseX, _
                                                            lastButtonUp.MouseY, _
                                                            lastButtonUp.MouseAbsX, _
                                                            lastButtonUp.MouseAbsY, _
                                                            lastButtonUp.IsInMouseNorth, _
                                                            lastButtonUp.IsInMouseSouth, _
                                                            lastButtonUp.IsInMouseEast, _
                                                            lastButtonUp.IsInMouseWest))
        Else
            '何もしない。
            'Xb.App.Out("Monitor.CheckClickEvent_None-Click")
        End If
    End Sub


#Region "IDisposable Support"
    Private _disposedValue As Boolean ' 重複する呼び出しを検出するには

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me._disposedValue Then
            If disposing Then
                Me._delegate = Nothing

                Me._history.Clear()
                Me._history = Nothing
            End If

            ' フックを削除
            If (Not UnhookWindowsHookEx(Me._mouseHandle)) Then
                Xb.Util.Out("Xb.Ui.Monitor.Dispose: グローバルフックの解除に失敗しました。")
                Throw New ApplicationException("グローバルフックの解除に失敗しました。")
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

#Region "グローバルフック処理のミニマム実装サンプル"
'
' --- 以下、マウスグローバルフックのミニマム実装例。 ---
'
'Imports System.Windows.Forms
'Imports System.Runtime.InteropServices
'Partial Public Class MouseCaptureTest
'    Inherits Form
'    ''' <summary>
'    ''' アプリケーション定義のフックプロシージャ"lpfn"をフックチェーン内にインストールします。
'    ''' </summary>
'    ''' <param name="idHook">フックタイプ</param>
'    ''' <param name="lpfn">フックプロシージャ</param>
'    ''' <param name="hInstance">アプリケーションインスタンスのハンドル</param>
'    ''' <param name="threadId">スレッドの識別子</param>
'    ''' <returns>フックプロシージャのハンドル</returns>
'    ''' <remarks>
'    ''' 
'    ''' SetWindowsHookEx
'    ''' http://msdn.microsoft.com/ja-jp/library/cc430103.aspx
'    ''' 
'    ''' マウス対象のグローバルフックなので、idHook = WH_MOUSE_LL(=14) になる。
'    ''' 
'    ''' フック時のコールバック関数が LowLevelMouseProc(=HookProcedureDelegate)。
'    ''' http://msdn.microsoft.com/ja-jp/library/cc430021.aspx
'    ''' 
'    ''' </remarks>
'    <DllImport("user32.dll", CharSet:=CharSet.Auto, _
'    CallingConvention:=CallingConvention.StdCall)> _
'    Public Shared Function SetWindowsHookEx(ByVal idHook As Integer, _
'                                            ByVal lpfn As HookProcedureDelegate, _
'                                            ByVal hInstance As IntPtr, _
'                                            ByVal threadId As Integer) As IntPtr
'    End Function


'    ''' <summary>
'    ''' SetWindowsHookEx 関数を使ってフックチェーン内にインストールされたフックプロシージャを削除します。
'    ''' </summary>
'    ''' <param name="idHook">削除対象のフックプロシージャのハンドル</param>
'    ''' <returns>関数が成功すると 0 以外の値、失敗すると 0 が返ります。</returns>
'    ''' <remarks>
'    ''' UnhookWindowsHookEx
'    ''' http://msdn.microsoft.com/ja-jp/library/cc430120.aspx
'    ''' </remarks>
'    <DllImport("user32.dll", CharSet:=CharSet.Auto, _
'    CallingConvention:=CallingConvention.StdCall)> _
'    Public Shared Function UnhookWindowsHookEx(ByVal idHook As IntPtr) As Boolean
'    End Function


'    ''' <summary>
'    ''' 現在のフックチェーン内の次のフックプロシージャに、フック情報を渡します。
'    ''' </summary>
'    ''' <param name="idHook">現在のフックのハンドル</param>
'    ''' <param name="nCode">フックプロシージャに渡すフックコード</param>
'    ''' <param name="wParam">フックプロシージャに渡す値</param>
'    ''' <param name="lParam">フックプロシージャに渡す値</param>
'    ''' <returns>フックチェーン内の次のフックプロシージャの戻り値</returns>
'    ''' <remarks>
'    ''' CallNextHookEx
'    ''' http://msdn.microsoft.com/ja-jp/library/cc429591.aspx
'    ''' </remarks>
'    <DllImport("user32.dll", CharSet:=CharSet.Auto, _
'    CallingConvention:=CallingConvention.StdCall)> _
'    Public Shared Function CallNextHookEx(ByVal idHook As IntPtr, _
'                                          ByVal nCode As Integer, _
'                                          ByVal wParam As IntPtr, _
'                                          ByVal lParam As IntPtr) As IntPtr
'    End Function

'    ''' <summary>
'    ''' 呼び出し側プロセスのアドレス空間に該当ファイルがマップされている場合、
'    ''' 指定されたモジュール名のモジュールハンドルを返します。
'    ''' </summary>
'    ''' <param name="lpModuleName">モジュール名</param>
'    ''' <returns>関数が成功時はモジュールのハンドル、失敗時はNullが返る。</returns>
'    ''' <remarks>
'    ''' GetModuleHandle
'    ''' http://msdn.microsoft.com/ja-jp/library/cc429129.aspx
'    ''' </remarks>
'    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
'    Public Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
'    End Function


'    ''' <summary>
'    ''' LowLevelMouseProc関数を示すデリゲート
'    ''' </summary>
'    ''' <param name="nCode">フックコード</param>
'    ''' <param name="wParam">
'    ''' メッセージ識別子
'    '''   WM_LBUTTONDOWN   : 0x0201
'    '''   WM_LBUTTONUP     : 0x0202
'    '''   WM_MOUSEMOVE     : 0x0200
'    '''   WM_MOUSEWHEEL    : 0x020A
'    '''   WM_MOUSEHWHEEL   : 0x020E
'    '''   WM_RBUTTONDOWN   : 0x0204
'    '''   WM_RBUTTONUP     : 0x0205
'    ''' </param>
'    ''' <param name="lParam">メッセージデータ(=MSLLHOOKSTRUCT構造体)</param>
'    ''' <returns>(強く推奨)フックチェーン内の次のフックプロシージャへの戻り値</returns>
'    ''' <remarks></remarks>
'    Public Delegate Function HookProcedureDelegate(nCode As Integer, _
'                                                   wParam As IntPtr, _
'                                                   lParam As IntPtr) As IntPtr


'    ''' <summary>
'    ''' 低レベルマウス入力イベント情報構造体
'    ''' </summary>
'    ''' <remarks>
'    ''' MSLLHOOKSTRUCT structure
'    ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms644970%28v=vs.85%29.aspx
'    ''' </remarks>
'    <StructLayout(LayoutKind.Sequential)> _
'    Public Class MouseHookStruct

'        ''' <summary>
'        ''' スクリーン上のX, Y座標
'        ''' </summary>
'        ''' <remarks></remarks>
'        Public point As PointStruct

'        ''' <summary>
'        ''' マウス情報値
'        ''' </summary>
'        ''' <remarks>
'        ''' WM_MOUSEWHEEL      : 0x020A  : マウスホイールが回った。
'        ''' WM_XBUTTONDOWN     : 0x020B  : マウスの "X"ボタン が押された。 "X"って何だ？
'        ''' WM_XBUTTONUP       : 0x020C  : マウスの "X"ボタン が離された。
'        ''' WM_XBUTTONDBLCLK   : 0x020D  : マウスの "X"ボタン がダブルクリックされた。
'        ''' WM_NCXBUTTONDOWN   : 0x00AB  : マウスの "X"ボタン が押された。 XBUTTONDOWN との違いが分からん。
'        ''' WM_NCXBUTTONUP     : 0x00AC  : 
'        ''' WM_NCXBUTTONDBLCLK : 0x00AD  : 
'        ''' </remarks>
'        Public mouseData As UInteger

'        ''' <summary>
'        ''' イベント割り込みフラグ
'        ''' </summary>
'        ''' <remarks></remarks>
'        Public flags As UInteger

'        ''' <summary>
'        ''' タイムスタンプ
'        ''' </summary>
'        ''' <remarks></remarks>
'        Public time As UInteger

'        ''' <summary>
'        ''' その他追加メッセージ
'        ''' </summary>
'        ''' <remarks></remarks>
'        Public dwExtraInfo As IntPtr
'    End Class


'    ''' <summary>
'    ''' X, Y座標を示す構造体
'    ''' </summary>
'    ''' <remarks>
'    ''' POINT structure
'    ''' http://msdn.microsoft.com/en-us/library/windows/desktop/dd162805%28v=vs.85%29.aspx
'    ''' </remarks>
'    <StructLayout(LayoutKind.Sequential)> _
'    Public Structure PointStruct
'        Public x As Integer
'        Public y As Integer
'    End Structure


'    'フックするメッセージタイプ、低レベルマウスイベントを示す。SetWindowsHookExで定義される。
'    'http://msdn.microsoft.com/ja-jp/library/cc430103.aspx
'    Private Const WH_MOUSE_LL As Integer = 14



'    Private _mouseHandle As IntPtr = IntPtr.Zero


'    ''' <summary>
'    ''' コンストラクタ
'    ''' </summary>
'    Public Sub New()
'        InitializeComponent()
'    End Sub


'    ''' <summary>
'    ''' ボタンクリック
'    ''' </summary>
'    ''' <param name="sender"></param>
'    ''' <param name="e"></param>
'    ''' <remarks></remarks>
'    Private Sub Button1_Click(ByVal sender As System.Object, _
'                              ByVal e As EventArgs) Handles Button1.Click

'        If (Me._mouseHandle = IntPtr.Zero) Then
'            'マウスのイベント取得を開始する。
'            Using processes As Process = Process.GetCurrentProcess()
'                Using mainModule As ProcessModule = processes.MainModule
'                    Me._mouseHandle = SetWindowsHookEx(WH_MOUSE_LL, _
'                                                    New HookProcedureDelegate(AddressOf OnMouseChanged), _
'                                                    GetModuleHandle(mainModule.ModuleName), 0)
'                End Using
'            End Using

'            If (Me._mouseHandle = IntPtr.Zero) Then
'                Console.WriteLine("SetWindowsHookEx Failed.")
'            End If
'        Else
'            ' フックを削除
'            If (UnhookWindowsHookEx(Me._mouseHandle)) Then
'                Me._mouseHandle = IntPtr.Zero
'            Else
'                Console.WriteLine("UnhookWindowsHookEx Failed.")
'            End If
'        End If

'    End Sub


'    ''' <summary>
'    ''' マウスイベント取得時のイベント処理
'    ''' </summary>
'    ''' <param name="nCode">フックコード</param>
'    ''' <param name="wParam">メッセージ識別子</param>
'    ''' <param name="lParam">メッセージデータ</param>
'    ''' <returns>(強く推奨)フックチェーン内の次のフックプロシージャへの戻り値</returns>
'    ''' <remarks></remarks>
'    Public Function OnMouseChanged(ByVal nCode As Integer, _
'                                   ByVal wParam As IntPtr, _
'                                   ByVal lParam As IntPtr) As IntPtr

'        If (nCode >= 0) Then

'            'アンマネージ領域からメッセージデータ構造体をマーシャリングして取得する。
'            Dim values As MouseHookStruct = DirectCast( _
'                Marshal.PtrToStructure(lParam, GetType(MouseHookStruct)),  _
'                MouseHookStruct _
'            )

'            Me.Text = "x = " & values.point.x.ToString() _
'                        & " : y = " & values.point.y.ToString() _
'                        & "  / " & Xb.Type.Byte.GetBitString(wParam.ToInt32())
'        End If

'        'フックチェーン内の次のフックプロシージャへ現在のフックを送る。
'        '下記を書かないでおくと、この処理以降のマウスフック待機中プロシージャがフックを受け取れない。
'        Return CallNextHookEx(Me._mouseHandle, nCode, wParam, lParam)

'    End Function

'End Class
#End Region