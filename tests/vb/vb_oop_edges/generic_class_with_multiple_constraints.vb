' vybe-test: vb/vb_oop_edges/generic_class_with_multiple_constraints
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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

Class Factory(Of T As {Class, New})
    Public Function Create() As T
        Return New T()
    End Function
End Class

Class Item
    Public Sub New()
        __Check(CStr("Item"), "Item")
    End Sub
End Class

Module M
    Sub Main()
        Dim f As New Factory(Of Item)()
        f.Create()
    End Sub
End Module
