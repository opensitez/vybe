' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_rank_length
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
        Dim arr(2, 3) As String
        __Check(CStr(arr.Rank), "2")
        __Check(CStr(arr.Length), "12")
        __Check(CStr(arr.GetLength(0)), "3")
        __Check(CStr(arr.GetLength(1)), "4")
    End Sub
End Module
