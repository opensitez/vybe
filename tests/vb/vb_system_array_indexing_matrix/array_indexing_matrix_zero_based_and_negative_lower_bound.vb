' vybe-test: vb/vb_system_array_indexing_matrix/array_indexing_matrix_zero_based_and_negative_lower_bound
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
        Dim values(-2 To 2) As Integer
        values(-2) = -2
        values(-1) = -1
        values(0) = 0
        values(1) = 1
        values(2) = 2

        __Check(CStr(values.GetLowerBound(0)), "-2")
        __Check(CStr(values.GetUpperBound(0)), "2")
        __Check(CStr(values(1) + values(-1)), "0")
    End Sub
End Module
