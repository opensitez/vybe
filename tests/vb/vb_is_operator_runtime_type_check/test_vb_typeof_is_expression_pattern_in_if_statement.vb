' vybe-test: vb/vb_is_operator_runtime_type_check/test_vb_typeof_is_expression_pattern_in_if_statement
' origin: languages/vb/tests/vb/test_vb_is_operator_runtime_type_check.rs

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

Module Program
    Sub Main()
        Dim item As Object = "Hello"
        If TypeOf item Is String Then
            Dim s As String = DirectCast(item, String)
            __Check(CStr(s.Length), "5")
        End If
    End Sub
End Module
