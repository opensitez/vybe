' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_reference_types
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
        Dim names(1, 1) As String
        names(0, 0) = "Alice"
        names(1, 1) = "Bob"
        __Check(CStr(names(0, 0)), "Alice")
        __Check(CStr(names(0, 1) Is Nothing), "True")
        __Check(CStr(names(1, 1)), "Bob")
    End Sub
End Module
