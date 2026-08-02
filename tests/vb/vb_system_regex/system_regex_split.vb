' vybe-test: vb/vb_system_regex/system_regex_split
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
        Dim input As String = "a, b; c |d"
        Dim pattern As String = "[,;\|]\s*"
        
        Dim parts As String() = Regex.Split(input, pattern)
        __Check(CStr(parts.Length), "4")
        __Check(CStr(parts(0)), "a")
        __Check(CStr(parts(3)), "d")
    End Sub
End Module
