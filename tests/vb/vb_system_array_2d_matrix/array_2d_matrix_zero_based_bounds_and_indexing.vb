' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_zero_based_bounds_and_indexing
' origin: languages/vb/tests/vb/test_vb_system_array_2d_matrix.rs

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
        Dim grid(0 To 1, 0 To 1) As String
        grid(0, 0) = "a"
        grid(0, 1) = "b"
        grid(1, 0) = "c"
        grid(1, 1) = "d"

        __Check(CStr(grid.GetLowerBound(0)), "0")
        __Check(CStr(grid.GetLowerBound(1)), "0")
        __Check(CStr(grid(1, 0) & grid(0, 1)), "cb")
    End Sub
End Module
