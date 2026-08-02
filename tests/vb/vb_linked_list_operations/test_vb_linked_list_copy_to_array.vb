' vybe-test: vb/vb_linked_list_operations/test_vb_linked_list_copy_to_array
' origin: languages/vb/tests/vb/test_vb_linked_list_operations.rs

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
        Dim ll As New LinkedList(Of Integer)()
        ll.AddLast(10)
        ll.AddLast(20)
        Dim arr(1) As Integer
        ll.CopyTo(arr, 0)
        __Check(CStr(String.Join(",", arr)), "10,20")
    End Sub
End Module
