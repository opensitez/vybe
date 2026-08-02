' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_object_array_mixed_types
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
        Dim obj(1) As Object
        obj(0) = 42
        obj(1) = "Hello"
        ReDim Preserve obj(2)
        obj(2) = True
        __Check(CStr(obj(0).ToString() & "|" & obj(1).ToString() & "|" & obj(2).ToString()), "42|Hello|True")
    End Sub
End Module
