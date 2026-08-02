' vybe-test: vb/vb_partial_methods_impl/partial_methods_unimplemented
' origin: languages/vb/tests/vb/test_vb_partial_methods_impl.rs

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

Partial Class Helper
    Partial Private Sub Log(msg As String)
    End Sub
    
    Public Sub DoWork()
        __Check(CStr("Start"), "Start")
        ' This call is compiled away if not implemented
        Log("Middle")
        __Check(CStr("End"), "End")
    End Sub
End Class

Module M
    Sub Main()
        Dim h As New Helper()
        h.DoWork()
    End Sub
End Module
