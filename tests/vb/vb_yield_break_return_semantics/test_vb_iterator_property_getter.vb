' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_property_getter
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

Class SequenceProvider
    Public ReadOnly Iterator Property Sequence As IEnumerable(Of Integer)
        Get
            Yield 100
            Yield 200
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New SequenceProvider()
        __Check(CStr(String.Join("+", p.Sequence)), "100+200")
    End Sub
End Module
