' vybe-test: vb/vb_system_string_api_matrix/string_api_matrix_replace_and_remove_slice
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
        Dim base As String = "a-b-c-d"
        __Check(CStr(base.Replace("-", "")), "abcd")
        __Check(CStr(base.Substring(2, 3)), "b-c")
        __Check(CStr(base.Remove(4)), "a-b-")
        __Check(CStr(base.Insert(0, "[")), "[a-b-c-d")
    End Sub
End Module
