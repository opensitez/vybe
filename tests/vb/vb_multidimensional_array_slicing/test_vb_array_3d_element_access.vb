' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_3d_element_access
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
        Dim cube(1, 1, 1) As Integer
        cube(0, 0, 0) = 10
        cube(1, 1, 1) = 99
        __Check(CStr(cube(0, 0, 0)), "10")
        __Check(CStr(cube(1, 1, 1)), "99")
    End Sub
End Module
