' vybe-test: vb/vb_list_get_range_insert_range/test_vb_list_as_read_only_wrapper
' origin: languages/vb/tests/vb/test_vb_list_get_range_insert_range.rs

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
        Dim list As New List(Of Integer) From {10, 20, 30}
        Dim ro As ReadOnlyCollection(Of Integer) = list.AsReadOnly()
        __Check(CStr(ro.Count & ":" & ro(0)), "3:10")
    End Sub
End Module
