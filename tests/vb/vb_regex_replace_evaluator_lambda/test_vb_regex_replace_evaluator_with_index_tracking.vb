' vybe-test: vb/vb_regex_replace_evaluator_lambda/test_vb_regex_replace_evaluator_with_index_tracking
' origin: languages/vb/tests/vb/test_vb_regex_replace_evaluator_lambda.rs

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

Module Program
    Sub Main()
        Dim input = "apple banana cherry"
        Dim index = 1
        Dim result = Regex.Replace(input, "\b\w+\b", Function(m)
            Dim res = index & "." & m.Value
            index += 1
            Return res
        End Function)
        __Check(CStr(result), "1.apple 2.banana 3.cherry")
    End Sub
End Module
