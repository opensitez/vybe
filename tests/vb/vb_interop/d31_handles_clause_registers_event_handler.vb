' vybe-test: vb/vb_interop/d31_handles_clause_registers_event_handler
' origin: languages/vb/tests/vb/vb_interop_test.rs

' ⛔ The extraction left this file holding a single token (`btn1`), which
' is not a VB program. Rewritten to what the name describes: a `Handles`
' clause binding a handler to a `WithEvents` field's event.

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


Class Button
    Public Event Click As EventHandler
    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Class Form
    Public WithEvents btn1 As New Button()

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        __P(CStr("btn1 clicked"))
    End Sub
End Class

Module Program
    Sub Main()
        Dim f As New Form()
        f.btn1.PerformClick()
        __Check("btn1 clicked")
    End Sub
End Module
