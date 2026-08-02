' vybe-test: vb/vb_linked_list_operations/test_vb_linked_list_add_before_add_after_node
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
        Dim ll As New LinkedList(Of String)()
        Dim node As LinkedListNode(Of String) = ll.AddLast("Target")
        ll.AddBefore(node, "Before")
        ll.AddAfter(node, "After")
        __Check(CStr(String.Join(",", ll)), "Before,Target,After")
    End Sub
End Module
