' vybe-test: vb/vb_list_sort_custom_comparer/test_vb_list_remove_all_predicate
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
        Dim list As New List(Of Integer) From {1, 2, 3, 4, 5, 6}
        Dim count As Integer = list.RemoveAll(Function(x) x Mod 2 = 0)
        __Check(CStr(count), "3")
        __Check(CStr(String.Join(",", list)), "1,3,5")
    End Sub
End Module
