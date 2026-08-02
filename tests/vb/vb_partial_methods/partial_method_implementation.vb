' vybe-test: vb/vb_partial_methods/partial_method_implementation
' origin: languages/vb/tests/vb/test_vb_partial_methods.rs

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

Partial Class Logger
    ' Declaration
    Partial Private Sub LogMessage(msg As String)
    End Sub
    
    Public Sub DoWork()
        __Check(CStr("Working"), "Working")
        LogMessage("Work completed")
    End Sub
End Class

Partial Class Logger
    ' Implementation
    Private Sub LogMessage(msg As String)
        __Check(CStr("LOG: " & msg), "LOG: Work completed")
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New Logger()
        l.DoWork()
    End Sub
End Module
