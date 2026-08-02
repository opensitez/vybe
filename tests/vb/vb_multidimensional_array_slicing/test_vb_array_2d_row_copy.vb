' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_row_copy
' origin: languages/vb/tests/vb/test_vb_multidimensional_array_slicing.rs

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
        Dim matrix(,) As Integer = {{10, 20, 30}, {40, 50, 60}}
        Dim row1(2) As Integer
        Buffer.BlockCopy(matrix, 0, row1, 0, 3 * sizeof(Integer))
        __Check(CStr(String.Join(",", row1)), "10,20,30")
    End Sub
End Module
