' vybe-test: vb/vb_custom_event_add_remove_raise/test_vb_custom_event_block_explicit_accessors
' origin: languages/vb/tests/vb/test_vb_custom_event_add_remove_raise.rs

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


Public Delegate Sub CustomHandler(msg As String)

Class Publisher
    Private _handlers As CustomHandler

    Public Custom Event StateChanged As CustomHandler
    AddHandler(value As CustomHandler)
    __P(CStr("Added"))
    _handlers = CType([Delegate].Combine(_handlers, value), CustomHandler)
End AddHandler

RemoveHandler(value As CustomHandler)
__P(CStr("Removed"))
_handlers = CType([Delegate].Remove(_handlers, value), CustomHandler)
End RemoveHandler

RaiseEvent(msg As String)
__P(CStr("Raising"))
_handlers?.Invoke(msg)
End RaiseEvent
End Event

Public Sub Trigger(msg As String)
    RaiseEvent StateChanged(msg)
End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New Publisher()
        Dim h As CustomHandler = Sub(m) __P(CStr("Received: " & m))
        AddHandler pub.StateChanged, h
        pub.Trigger("Data1")
        RemoveHandler pub.StateChanged, h
        __Check("Added
Raising
Received: Data1
Removed")
    End Sub
End Module
