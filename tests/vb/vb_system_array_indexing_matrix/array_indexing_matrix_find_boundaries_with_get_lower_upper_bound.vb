' vybe-test: vb/vb_system_array_indexing_matrix/array_indexing_matrix_find_boundaries_with_get_lower_upper_bound
' origin: languages/vb/tests/vb/test_vb_system_array_indexing_matrix.rs

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
        Dim values() As Integer = {10, 20, 30}

        __Check(CStr(values.GetLowerBound(0)), "0")
        __Check(CStr(values.GetUpperBound(0)), "2")
        __Check(CStr(values(values.GetLowerBound(0))), "10")
        __Check(CStr(values(values.GetUpperBound(0))), "30")
    End Module
