' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_handler_multithreaded_subscription
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
Imports System.Threading
Imports System.Threading.Tasks
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


Class ConcurrentNotifier
    Private lockObj As New Object()
    Private multicast As EventHandler

    Public Custom Event SharedEvent As EventHandler
        AddHandler(value As EventHandler)
            SyncLock lockObj
                multicast = CType(Delegate.Combine(multicast, value), EventHandler)
            End SyncLock
        End AddHandler
        RemoveHandler(value As EventHandler)
            SyncLock lockObj
                multicast = CType(Delegate.Remove(multicast, value), EventHandler)
            End SyncLock
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Dim copy As EventHandler
            SyncLock lockObj
                copy = multicast
            End SyncLock
            If copy IsNot Nothing Then copy(sender, e)
        End RaiseEvent
    End Event

    Public Function GetCount() As Integer
        SyncLock lockObj
            Return If(multicast IsNot Nothing, multicast.GetInvocationList().Length, 0)
        End SyncLock
    End Function
End Class

Module Program
    Sub Main()
        Dim cn As New ConcurrentNotifier()
        Dim tasks(3) As Task
        For i As Integer = 0 To 3
            tasks(i) = Task.Run(Sub()
                AddHandler cn.SharedEvent, Sub(s, e)
                    __Check("Concurrent Handlers: 4")
                End Sub
            End Sub)
        Next
        Task.WaitAll(tasks)
        __P(CStr("Concurrent Handlers: " & cn.GetCount()))
    End Sub
End Module
