' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_construct_and_read
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
        Dim grid(1 To 2, 1 To 3) As Integer
        grid(1, 1) = 10
        grid(1, 2) = 20
        grid(1, 3) = 30
        grid(2, 1) = 40
        grid(2, 2) = 50
        grid(2, 3) = 60

        __Check(CStr(grid.GetLength(0)), "2")
        __Check(CStr(grid.GetLength(1)), "3")
        __Check(CStr(grid(1, 1) + grid(1, 2) + grid(2, 3)), "90")
    End Sub
End Module
