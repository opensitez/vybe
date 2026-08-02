' vybe-test: vb/vb_list_sort_custom_comparer/test_vb_list_find_index_matching
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

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"one", "two", "three", "four"}
        Dim idx As Integer = list.FindIndex(Function(s) s.StartsWith("t"))
        __Check(CStr(idx), "1")
    End Sub
End Module
