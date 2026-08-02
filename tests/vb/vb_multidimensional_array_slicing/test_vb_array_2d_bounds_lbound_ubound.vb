' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_bounds_lbound_ubound
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
        Dim arr(2, 4) As Integer
        __Check(CStr(arr.GetLowerBound(0)), "0")
        __Check(CStr(arr.GetUpperBound(0)), "2")
        __Check(CStr(arr.GetLowerBound(1)), "0")
        __Check(CStr(arr.GetUpperBound(1)), "4")
    End Sub
End Module
