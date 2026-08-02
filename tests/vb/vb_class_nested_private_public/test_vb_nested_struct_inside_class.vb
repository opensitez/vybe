' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_struct_inside_class
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

Class Graph
    Public Structure Node
        Public ID As Integer
        Public Label As String
        Public Sub New(id As Integer, label As String)
            Me.ID = id : Me.Label = label
        End Sub
    End Structure
End Class

Module Program
    Sub Main()
        Dim n As New Graph.Node(1, "Root")
        __Check(CStr(n.ID & ":" & n.Label), "1:Root")
    End Sub
End Module
