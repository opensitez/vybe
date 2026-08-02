' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_4d_bounds
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
        Dim arr(1, 2, 3, 4) As Double
        __Check(CStr(arr.Rank), "4")
        __Check(CStr(arr.GetLength(3)), "5")
    End Sub
End Module
