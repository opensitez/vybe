' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_multidimensional_last_dimension_only
' origin: languages/vb/tests/vb/test_vb_array_resize_preserve_semantics.rs

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
        Dim mat(1, 1) As Integer
        mat(0, 0) = 1 : mat(0, 1) = 2
        mat(1, 0) = 3 : mat(1, 1) = 4

        ReDim Preserve mat(1, 2)
        __Check(CStr(mat(0, 0) & "," & mat(0, 1) & "," & mat(0, 2)), "1,2,0")
        __Check(CStr(mat(1, 0) & "," & mat(1, 1) & "," & mat(1, 2)), "3,4,0")
    End Sub
End Module
