' vybe-test: vb/vb_regex_replace_evaluator_lambda/test_vb_regex_replace_evaluator_conditional_replacement
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
        Dim input = "10 25 50 100"
        Dim result = Regex.Replace(input, "\d+", Function(m)
            Dim val = Integer.Parse(m.Value)
            If val > 30 Then Return "HIGH" Else Return "LOW"
        End Function)
        __Check(CStr(result), "LOW LOW HIGH HIGH")
    End Sub
End Module
