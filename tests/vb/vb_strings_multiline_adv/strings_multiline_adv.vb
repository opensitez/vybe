' vybe-test: vb/vb_strings_multiline_adv/strings_multiline_adv
' origin: languages/vb/tests/vb/test_vb_strings_multiline_adv.rs

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
        ' Multiline strings preserve whitespace
        Dim s As String = "Line 1
    Line 2
Line 3"
        Dim lines = s.Split({vbLf}, StringSplitOptions.None)
        __Check(CStr(lines.Length), "3")
        __Check(CStr(lines(1).TrimStart()), "Line 2")
    End Sub
End Module
