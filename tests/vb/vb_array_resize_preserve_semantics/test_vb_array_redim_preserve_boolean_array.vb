' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_boolean_array
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
        Dim flags(1) As Boolean
        flags(0) = True
        ReDim Preserve flags(3)
        __Check(CStr(flags(0) & "," & flags(1) & "," & flags(2)), "True,False,False")
    End Sub
End Module
