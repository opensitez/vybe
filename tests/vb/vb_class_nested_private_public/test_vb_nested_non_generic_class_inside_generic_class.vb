' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_non_generic_class_inside_generic_class
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Class OuterList(Of T)
    Public Class Node
        Public Element As T
        Public NextNode As Node
        Public Sub New(e As T)
            Element = e
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Dim node1 As New OuterList(Of Integer).Node(10)
        Dim node2 As New OuterList(Of Integer).Node(20)
        node1.NextNode = node2
        __Check(CStr(node1.Element & "->" & node1.NextNode.Element), "10->20")
    End Sub
End Module
