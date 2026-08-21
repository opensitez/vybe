' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_invocation_list_subscriber_count
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

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


Class Publisher
    Private delegateList As EventHandler

    Public Custom Event StatusChanged As EventHandler
        AddHandler(value As EventHandler)
            delegateList = CType(Delegate.Combine(delegateList, value), EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
            delegateList = CType(Delegate.Remove(delegateList, value), EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            If delegateList IsNot Nothing Then delegateList(sender, e)
        End RaiseEvent
    End Event

    Public Function GetSubscriberCount() As Integer
        Return If(delegateList IsNot Nothing, delegateList.GetInvocationList().Length, 0)
    End Function
End Class

Module Program
    Sub Main()
        Dim p As New Publisher()
        Dim h1 As EventHandler = Sub(s, e) __P(CStr("H1"))
        Dim h2 As EventHandler = Sub(s, e) __P(CStr("H2"))

        AddHandler p.StatusChanged, h1
        AddHandler p.StatusChanged, h2
        __P(CStr("Subscribers: " & p.GetSubscriberCount()))
        __Check("Subscribers: 2")
    End Sub
End Module
