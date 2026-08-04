' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_custom_event_thread_sync_lock
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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

Class ThreadSafeEventPublisher
    Private syncObj As New Object()
    Private internalDelegate As Action

    Public Custom Event SecureEvent As Action
        AddHandler(value As Action)
            SyncLock syncObj
                internalDelegate = CType([Delegate].Combine(internalDelegate, value), Action)
            End SyncLock
        End AddHandler
        RemoveHandler(value As Action)
            SyncLock syncObj
                internalDelegate = CType([Delegate].Remove(internalDelegate, value), Action)
            End SyncLock
        End RemoveHandler
        RaiseEvent()
            Dim copy As Action = Nothing
            SyncLock syncObj
                copy = internalDelegate
            End SyncLock
            If copy IsNot Nothing Then copy()
        End RaiseEvent
    End Event

    Public Sub Run()
        RaiseEvent SecureEvent()
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New ThreadSafeEventPublisher()
        AddHandler pub.SecureEvent, Sub() __P(CStr("ThreadSafe Event Fired"))
        pub.Run()
        __Check("ThreadSafe Event Fired")
    End Sub
End Module
