' vybe-test: vb/vb_array_apis/array_exists_reports_predicate_match
' origin: languages/vb/tests/vb/test_vb_array_apis.rs

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
        Dim values As Integer() = {1, 3, 5}
        __Check(CStr(Array.Exists(values, Function(value As Integer) value = 3)), "True")
    End Sub
End Module
