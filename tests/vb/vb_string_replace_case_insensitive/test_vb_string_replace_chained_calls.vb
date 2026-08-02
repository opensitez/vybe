' vybe-test: vb/vb_string_replace_case_insensitive/test_vb_string_replace_chained_calls
' origin: languages/vb/tests/vb/test_vb_string_replace_case_insensitive.rs

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
        Dim s As String = "a-b-c-d"
        Dim result = s.Replace("a", "1").Replace("b", "2").Replace("c", "3").Replace("d", "4")
        __Check(CStr(result), "1-2-3-4")
    End Sub
End Module
