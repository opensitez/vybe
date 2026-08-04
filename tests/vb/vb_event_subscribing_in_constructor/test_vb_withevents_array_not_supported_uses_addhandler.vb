' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_array_not_supported_uses_addhandler
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

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

Class Button
    Public Property ID As Integer
    Public Event Click As EventHandler
    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Class FormContainer
    Private buttons As Button()

    Public Sub New()
        buttons = New Button() {New Button With {.ID = 1}, New Button With {.ID = 2}}
        For Each btn In buttons
            AddHandler btn.Click, AddressOf OnButtonClick
        Next
    End Sub

    Private Sub OnButtonClick(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        __P(CStr("Button " & btn.ID & " Clicked"))
    End Sub

    Public Sub TestClicks()
        buttons(0).PerformClick()
        buttons(1).PerformClick()
    End Sub
End Class

Module Program
    Sub Main()
        Dim form As New FormContainer()
        form.TestClicks()
        __Check("Button 1 Clicked
Button 2 Clicked")
    End Sub
End Module
