' vybe-test: vb/vb_regex_replace_evaluator_lambda/test_vb_regex_replace_mask_sensitive_data
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
        Dim cc = "1234-5678-9012-3456"
        Dim masked = Regex.Replace(cc, "\d{4}-\d{4}-\d{4}-", "****-****-****-")
        __Check(CStr(masked), "****-****-****-3456")
    End Sub
End Module
