' vybe-test: vb/vb_system_string_api_matrix/string_api_matrix_formatting_contracts
' origin: languages/vb/tests/vb/test_vb_system_string_api_matrix.rs

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

Module M
    Sub Main()
        Dim a As String = String.Format("A={0}, B={1}", 1, "x")
        Dim b As String = String.Format("{0,5}", 42)
        Dim c As String = String.Format("{0:F1}", 3.14159)
        __Check(CStr(a), "A=1, B=x")
        __Check(CStr(b), "   42")
        __Check(CStr(c), "3.1")
    End Sub
End Module
