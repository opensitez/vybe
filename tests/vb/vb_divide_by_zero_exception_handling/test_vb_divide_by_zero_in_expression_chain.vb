' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_divide_by_zero_in_expression_chain
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

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
        Try
            Dim val As Integer = (100 + 50) \ (10 - 10)
        Catch ex As DivideByZeroException
            __Check(CStr("Chain DivideByZeroException Handled"), "Chain DivideByZeroException Handled")
        End Try
    End Sub
End Module
