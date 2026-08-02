' vybe-test: vb/vb_string_padleft_padright_formatting/test_vb_string_pad_combined_center_alignment
' origin: languages/vb/tests/vb/test_vb_string_padleft_padright_formatting.rs

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
        Dim s As String = "Title"
        ' Center in width 11: 3 spaces left, 3 spaces right
        Dim centered As String = s.PadLeft(8).PadRight(11)
        __Check(CStr("'" & centered & "'"), "'   Title   '")
    End Sub
End Module
