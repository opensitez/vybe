' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_jagged_to_rectangular_distinction
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
        Dim rect(1, 1) As Integer
        rect(0, 0) = 1
        rect(0, 1) = 2
        rect(1, 0) = 3
        rect(1, 1) = 4

        Dim rows()() As Integer = {New Integer() {1, 2}, New Integer() {3, 4, 5}}
        Dim totalRect As Integer = rect(0, 0) + rect(0, 1) + rect(1, 0) + rect(1, 1)
        Dim totalJagged As Integer = rows(0)(0) + rows(0)(1) + rows(1)(0) + rows(1)(1)

        __Check(CStr(totalRect), "10")
        __Check(CStr(totalJagged), "10")
        __Check(CStr(rows(1).Length), "3")
    End Sub
End Module
