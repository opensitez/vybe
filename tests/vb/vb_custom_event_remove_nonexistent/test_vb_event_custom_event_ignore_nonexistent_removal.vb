' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_custom_event_ignore_nonexistent_removal
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

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


Class CustomManager
    Private handlers As Action

    Public Custom Event Work As Action
    AddHandler(value As Action)
    handlers = CType([Delegate].Combine(handlers, value), Action)
End AddHandler
RemoveHandler(value As Action)
Dim newHandlers = CType([Delegate].Remove(handlers, value), Action)
If newHandlers Is Nothing AndAlso handlers IsNot Nothing Then
    __P(CStr("All Handlers Removed"))
End If
handlers = newHandlers
End RemoveHandler
RaiseEvent()
If handlers IsNot Nothing Then handlers()
End RaiseEvent
End Event

Public Sub Run()
    RaiseEvent Work()
End Sub
End Class

Module Program
    Private Sub Sub1()
        __P(CStr("Sub1"))
    End Sub

    Sub Main()
        Dim cm As New CustomManager()
        AddHandler cm.Work, AddressOf Sub1
        RemoveHandler cm.Work, AddressOf Sub1
        __Check("All Handlers Removed")
    End Sub
End Module
