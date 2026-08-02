' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_empty_generator
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
    Private Iterator Function EmptyGen() As IEnumerable(Of String)
        If False Then Yield "Never"
    End Function

    Sub Main()
        Dim list As New List(Of String)(EmptyGen())
        __Check(CStr(list.Count), "0")
    End Sub
End Module
