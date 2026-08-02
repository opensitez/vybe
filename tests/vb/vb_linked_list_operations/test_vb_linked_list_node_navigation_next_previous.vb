' vybe-test: vb/vb_linked_list_operations/test_vb_linked_list_node_navigation_next_previous
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
        ll.AddLast(30)

        Dim node As LinkedListNode(Of Integer) = ll.First.Next
        __Check(CStr(node.Value), "20")
        __Check(CStr(node.Previous.Value), "10")
        __Check(CStr(node.Next.Value), "30")
    End Sub
End Module
