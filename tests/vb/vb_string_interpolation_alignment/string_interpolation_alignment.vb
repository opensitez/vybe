' vybe-test: vb/vb_string_interpolation_alignment/string_interpolation_alignment
' origin: languages/vb/tests/vb/test_vb_string_interpolation_alignment.rs

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
        Dim val = 42
        
        ' String interpolation with alignment
        Dim s = $"[{val,5}]"
        __Check(CStr(s), "[   42]")
    End Sub
End Module
