' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_shrink_1d
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
        Dim arr(4) As String
        arr(0) = "One"
        arr(1) = "Two"
        arr(2) = "Three"
        arr(3) = "Four"
        arr(4) = "Five"

        ReDim Preserve arr(1)
        __Check(CStr(arr.Length), "2")
        __Check(CStr(arr(0) & "," & arr(1)), "One,Two")
    End Sub
End Module
