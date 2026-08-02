' vybe-test: vb/vb_interop/a04_method_with_params_called_from_another_method
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Formatter
    Public Function Wrap(s As String, prefix As String, suffix As String) As String
        Return prefix & s & suffix
    End Function
    Public Function WrapBrackets(s As String) As String
        Return Wrap(s, "[", "]")
    End Function
End Class
Dim f As New Formatter()
__Check(CStr(f.WrapBrackets("hello")), "[hello]")
