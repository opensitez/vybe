' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_multiple_enumerators_independent_state
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
    Private Iterator Function CounterGen() As IEnumerable(Of Integer)
        Yield 1
        Yield 2
    End Function

    Sub Main()
        Dim enumerable = CounterGen()
        Dim e1 = enumerable.GetEnumerator()
        Dim e2 = enumerable.GetEnumerator()

        e1.MoveNext()
        __Check(CStr(e1.Current & "|" & e2.MoveNext() & "|" & e2.Current), "1|True|1")
    End Sub
End Module
