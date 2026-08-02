' vybe-test: vb/vb_string_instr/string_instr
' origin: languages/vb/tests/vb/test_vb_string_instr.rs

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
        Dim s As String = "abracadabra"
        
        ' InStr (start_pos, string1, string2)
        ' 1-based index returns
        __Check(CStr(InStr(1, s, "a")), "1")
        __Check(CStr(InStr(2, s, "a")), "4")
        
        ' InStrRev searches from right to left
        __Check(CStr(InStrRev(s, "a")), "11")
    End Sub
End Module
