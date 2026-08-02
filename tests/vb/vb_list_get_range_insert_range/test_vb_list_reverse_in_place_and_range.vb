' vybe-test: vb/vb_list_get_range_insert_range/test_vb_list_reverse_in_place_and_range
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

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3, 4, 5}
        list.Reverse(1, 3) ' Reverse elements from index 1 (len 3): 2,3,4 -> 4,3,2
        __Check(CStr(String.Join(",", list)), "1,4,3,2,5")
    End Sub
End Module
