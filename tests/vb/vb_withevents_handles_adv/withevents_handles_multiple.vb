' vybe-test: vb/vb_withevents_handles_adv/withevents_handles_multiple
' origin: languages/vb/tests/vb/test_vb_withevents_handles_adv.rs

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

Class Button
    Public Event Click()
    Public Sub PerformClick()
        RaiseEvent Click()
    End Sub
End Class

Class Form
    Private WithEvents btn1 As New Button()
    Private WithEvents btn2 As New Button()
    
    ' Handles multiple events
    Private Sub CommonClickHandler() Handles btn1.Click, btn2.Click
        __P(CStr("Clicked"))
    End Sub
    
    Public Sub Test()
        btn1.PerformClick()
        btn2.PerformClick()
    End Sub
End Class

Module M
    Sub Main()
        Dim f As New Form()
        f.Test()
        __Check("Clicked
Clicked")
    End Sub
End Module
