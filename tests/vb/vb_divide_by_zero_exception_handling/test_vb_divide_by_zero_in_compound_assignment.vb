' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_divide_by_zero_in_compound_assignment
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
            Dim x As Integer = 100
            Dim zero As Integer = 0
            x \= zero
        Catch ex As DivideByZeroException
            __Check(CStr("Compound Backslash DivideByZeroException Handled"), "Compound Backslash DivideByZeroException Handled")
        End Try
    End Sub
End Module
