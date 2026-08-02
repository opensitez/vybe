' vybe-test: vb/vb_system_regex/system_regex_replace
' origin: languages/vb/tests/vb/test_vb_system_regex.rs

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

Imports System.Text.RegularExpressions

Module M
    Sub Main()
        Dim input As String = "apple banana apple"
        Dim pattern As String = "apple"
        
        Dim result As String = Regex.Replace(input, pattern, "orange")
        __Check(CStr(result), "orange banana orange")
    End Sub
End Module
