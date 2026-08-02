' vybe-test: vb/vb_regex_replace_evaluator_lambda/test_vb_regex_replace_with_count_limit
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
        Dim input = "a b a b a"
        Dim result = Regex.Replace(input, "a", "X", RegexOptions.None, TimeSpan.FromSeconds(1))
        Dim resultLimit = New Regex("a").Replace(input, "X", 2)
        __Check(CStr(resultLimit), "X b X b a")
    End Sub
End Module
