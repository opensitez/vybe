' vybe-test: vb/vb_list_remove_all_predicate/test_vb_list_remove_all_address_of_predicate
' origin: languages/vb/tests/vb/test_vb_list_remove_all_predicate.rs

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

Module Filters
    Public Function IsNegative(n As Integer) As Boolean
        Return n < 0
    End Function
End Module

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, -5, 20, -15, 30}
        list.RemoveAll(AddressOf Filters.IsNegative)
        __Check(CStr(String.Join(",", list)), "10,20,30")
    End Sub
End Module
