' vybe-test: vb/vb_explicit_conversions/explicit_trycast_failure_nothing
' origin: languages/vb/tests/vb/test_vb_explicit_conversions.rs

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

Class C
End Class
Module M
Sub Main()
Dim obj As Object = "Hello"
Dim c1 = TryCast(obj, C)
__Check(CStr(c1 Is Nothing), "True")
End Sub
End Module
