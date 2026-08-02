' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_conversion_from_value_tuple
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

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
        Dim rawTuple As ValueTuple(Of String, Integer) = ValueTuple.Create("Raw", 99)
        Dim namedTuple As (Tag As String, Val As Integer) = rawTuple
        __Check(CStr(namedTuple.Tag & "=" & namedTuple.Val), "Raw=99")
    End Sub
End Module
