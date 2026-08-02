' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_event_handler_relaxation_allowing_fewer_parameters
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Class EventPublisher
    Public Event Click As EventHandler
    Public Sub Fire()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    ' VB.NET allows delegate relaxation: omitting parameters
    Private Sub ParameterlessHandler()
        __Check(CStr("Parameterless Handler Invoked"), "Parameterless Handler Invoked")
    End Sub

    Sub Main()
        Dim ep As New EventPublisher()
        AddHandler ep.Click, AddressOf ParameterlessHandler
        ep.Fire()
    End Sub
End Module
