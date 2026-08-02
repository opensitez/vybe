' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_with_lambda_expression_property
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

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
        Dim obj = New With {
            .Multiplier = 2,
            .Calc = Function(n As Integer) n * 2
        }
        __Check(CStr(obj.Calc(10)), "20")
    End Sub
End Module
