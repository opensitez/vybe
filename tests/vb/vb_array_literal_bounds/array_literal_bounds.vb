' vybe-test: vb/vb_array_literal_bounds/array_literal_bounds
' origin: languages/vb/tests/vb/test_vb_array_literal_bounds.rs

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
        ' Array literal bounds inference for multi-dimensional arrays
        Dim arr(,) = {{1, 2}, {3, 4}, {5, 6}}
        
        __Check(CStr(arr.GetLength(0)), "3")
        __Check(CStr(arr.GetLength(1)), "2")
        __Check(CStr(arr(2, 1)), "6")
    End Sub
End Module
