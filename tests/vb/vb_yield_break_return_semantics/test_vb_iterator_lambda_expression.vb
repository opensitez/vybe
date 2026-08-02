' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_lambda_expression
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

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
Imports System.Collections.Generic

Module Program
    Sub Main()
        ' Iterator lambda expression syntax
        Dim gen As Func(Of IEnumerable(Of Integer)) = Iterator Function()
            Yield 5
            Yield 10
        End Function

        __Check(CStr(String.Join("+", gen())), "5+10")
    End Sub
End Module
