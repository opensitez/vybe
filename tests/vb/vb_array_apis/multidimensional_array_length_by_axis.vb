' vybe-test: vb/vb_array_apis/multidimensional_array_length_by_axis
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
        Dim matrix(1, 2) As Integer
        __Check(CStr(matrix.GetLength(0)), "2")
        __Check(CStr(matrix.GetLength(1)), "3")
    End Sub
End Module
