' vybe-test: vb/vb_arrays_bounds/array_declaration_upper_bound
' origin: languages/vb/tests/vb/test_vb_arrays_bounds.rs

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
        ' In VB, declaring an array with (5) means the upper bound is 5,
        ' so there are 6 elements (0 through 5).
        Dim arr(5) As Integer
        __Check(CStr(arr.Length), "6")
    End Sub
End Module
