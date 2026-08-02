' vybe-test: vb/vb_list_sort_custom_comparer/test_vb_list_as_read_only_wrapper
' origin: languages/vb/tests/vb/test_vb_list_sort_custom_comparer.rs

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
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"A", "B", "C"}
        Dim ro As ReadOnlyCollection(Of String) = list.AsReadOnly()
        __Check(CStr(ro.Count), "3")
        __Check(CStr(ro(1)), "B")
    End Sub
End Module
