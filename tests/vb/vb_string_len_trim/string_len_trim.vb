' vybe-test: vb/vb_string_len_trim/string_len_trim
' origin: languages/vb/tests/vb/test_vb_string_len_trim.rs

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
        Dim s As String = "  VB.NET  "
        
        ' Len measures string length (or variable byte size, but mostly string length)
        __Check(CStr(Len(s)), "10")
        
        ' Trim functions
        __Check(CStr("[" & Trim(s) & "]"), "[VB.NET]")
        __Check(CStr("[" & LTrim(s) & "]"), "[VB.NET  ]")
        __Check(CStr("[" & RTrim(s) & "]"), "[  VB.NET]")
    End Sub
End Module
