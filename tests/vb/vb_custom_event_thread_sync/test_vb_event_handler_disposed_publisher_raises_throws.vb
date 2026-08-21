' vybe-test: vb/vb_custom_event_thread_sync/test_vb_event_handler_disposed_publisher_raises_throws
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Imports System
Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module


Class DisposablePublisher
    Implements IDisposable
    Public Event DataEvent As EventHandler
    Private isDisposed As Boolean = False

    Public Sub Fire()
        If isDisposed Then Throw New ObjectDisposedException("DisposablePublisher")
        RaiseEvent DataEvent(Me, EventArgs.Empty)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        isDisposed = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim dp As New DisposablePublisher()
        dp.Dispose()
        Try
            dp.Fire()
        Catch ex As ObjectDisposedException
            __P(CStr("ObjectDisposedException Caught on Fire"))
        End Try
        __Check("ObjectDisposedException Caught on Fire")
    End Sub
End Module
