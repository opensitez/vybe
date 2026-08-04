' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_interlocked_exchange_accessor
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

Imports System
Imports System.Threading

Class InterlockedEventSource
    Private handlers As EventHandler

    Public Custom Event FastEvent As EventHandler
        AddHandler(value As EventHandler)
            Dim oldHandlers As EventHandler = Nothing
            Dim newHandlers As EventHandler = Nothing
            Do
                oldHandlers = handlers
                newHandlers = CType(Delegate.Combine(oldHandlers, value), EventHandler)
            Loop While Interlocked.CompareExchange(handlers, newHandlers, oldHandlers) IsNot oldHandlers
        End AddHandler

        RemoveHandler(value As EventHandler)
            Dim oldHandlers As EventHandler = Nothing
            Dim newHandlers As EventHandler = Nothing
            Do
                oldHandlers = handlers
                newHandlers = CType(Delegate.Remove(oldHandlers, value), EventHandler)
            Loop While Interlocked.CompareExchange(handlers, newHandlers, oldHandlers) IsNot oldHandlers
        End RemoveHandler

        RaiseEvent(sender As Object, e As EventArgs)
            Dim currentHandlers As EventHandler = Volatile.Read(handlers)
            If currentHandlers IsNot Nothing Then currentHandlers(sender, e)
        End RaiseEvent
    End Event

    Public Sub Fire()
        RaiseEvent FastEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ies As New InterlockedEventSource()
        AddHandler ies.FastEvent, Sub(s, e) __P(CStr("Interlocked Event Fired"))
        ies.Fire()
        __Check("Interlocked Event Fired")
    End Sub
End Module
