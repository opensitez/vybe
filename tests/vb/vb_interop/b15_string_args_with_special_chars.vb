' vybe-test: vb/vb_interop/b15_string_args_with_special_chars
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

Function Echo(s As String) As String
    Return s
End Function
__Check(CStr(Echo("hello world")), "hello world")
__Check(CStr(Echo("it's")), "it's")
__Check(CStr(Echo("a&b")), "a&b")
