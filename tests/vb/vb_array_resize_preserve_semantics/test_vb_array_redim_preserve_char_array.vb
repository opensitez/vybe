' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_char_array
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
        Dim chars(1) As Char
        chars(0) = "A"c
        chars(1) = "B"c
        ReDim Preserve chars(3)
        __Check(CStr(chars(0) & chars(1) & "|" & (chars(2) = ChrW(0))), "AB|True")
    End Sub
End Module
