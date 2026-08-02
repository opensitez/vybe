' vybe-test: vb/vb_sorted_list_operations/test_vb_sorted_list_capacity_trim_excess
' origin: languages/vb/tests/vb/test_vb_sorted_list_operations.rs

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
        Dim list As New SortedList(Of Integer, Integer)(100)
        list.Add(1, 10)
        list.TrimExcess()
        __Check(CStr(list.Capacity >= 1), "True")
    End Sub
End Module
