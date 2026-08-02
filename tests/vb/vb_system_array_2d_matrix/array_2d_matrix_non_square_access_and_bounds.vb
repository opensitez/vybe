' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_non_square_access_and_bounds
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
        Dim matrix(1 To 2, 1 To 4) As Integer
        __Check(CStr(matrix.Rank), "2")
        __Check(CStr(matrix.GetUpperBound(0)), "2")
        __Check(CStr(matrix.GetUpperBound(1)), "4")
    End Sub
End Module
