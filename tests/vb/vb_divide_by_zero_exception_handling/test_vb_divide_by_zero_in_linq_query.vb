' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_divide_by_zero_in_linq_query
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
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 0, 30}
        Try
            Dim results = (From n In numbers Select 100 \ n).ToList()
        Catch ex As DivideByZeroException
            __Check(CStr("LINQ DivideByZeroException Handled"), "LINQ DivideByZeroException Handled")
        End Try
    End Sub
End Module
