' vybe-test: vb/vb_partial_methods/partial_method_no_implementation
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

Partial Class SilencedLogger
    ' Declaration only, no implementation
    Partial Private Sub LogMessage(msg As String)
    End Sub
    
    Public Sub DoWork()
        __Check(CStr("Working"), "Working")
        ' If not implemented, call is removed entirely by compiler
        LogMessage("Work completed")
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New SilencedLogger()
        l.DoWork()
    End Sub
End Module
