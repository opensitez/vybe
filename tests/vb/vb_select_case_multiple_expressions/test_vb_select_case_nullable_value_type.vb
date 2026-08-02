' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_nullable_value_type
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

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

Imports System

Module Program
    Sub Main()
        Dim n As Integer? = 42
        Select Case n
            Case HasValue
                __Check(CStr("Value: " & n.Value), "Value: 42")
        End Select
    End Sub
End Module
