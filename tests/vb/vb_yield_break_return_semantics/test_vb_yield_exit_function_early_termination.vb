' vybe-test: vb/vb_yield_break_return_semantics/test_vb_yield_exit_function_early_termination
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

Imports System.Collections.Generic

Module Program
    Private Iterator Function GeneratorWithExit() As IEnumerable(Of String)
        Yield "A"
        If True Then Exit Function
        Yield "B"
    End Function

    Sub Main()
        Dim items = GeneratorWithExit()
        __Check(CStr(String.Join("", items)), "A")
    End Sub
End Module
