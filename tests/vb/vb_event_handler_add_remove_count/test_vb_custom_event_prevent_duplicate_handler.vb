' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_prevent_duplicate_handler
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
Imports System.Collections.Generic

Class UniquePublisher
    Private handlerList As New List(Of EventHandler)()

    Public Custom Event UniqueEvent As EventHandler
        AddHandler(value As EventHandler)
            If Not handlerList.Contains(value) Then handlerList.Add(value)
        End AddHandler
        RemoveHandler(value As EventHandler)
            handlerList.Remove(value)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            For Each h In handlerList
                h(sender, e)
            Next
        End RaiseEvent
    End Event

    Public Sub Trigger()
        RaiseEvent UniqueEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New UniquePublisher()
        Dim count = 0
        Dim handler As EventHandler = Sub(s, e) count += 1

        AddHandler p.UniqueEvent, handler
        AddHandler p.UniqueEvent, handler ' Duplicate add ignored by custom logic!
        p.Trigger()
        __P(CStr(count))
        __Check("1")
    End Sub
End Module
