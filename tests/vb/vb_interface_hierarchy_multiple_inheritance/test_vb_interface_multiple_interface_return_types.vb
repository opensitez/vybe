' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_multiple_interface_return_types
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IEntity
    ReadOnly Property ID As Integer
End Interface

Interface IAudit
    ReadOnly Property CreatedAt As String
End Interface

Class AuditedEntity
    Implements IEntity, IAudit
    Public ReadOnly Property ID As Integer Implements IEntity.ID
        Get
            Return 101
        End Get
    End Property
    Public ReadOnly Property CreatedAt As String Implements IAudit.CreatedAt
        Get
            Return "2025-01-01"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim ae As New AuditedEntity()
        __Check(CStr(ae.ID & "|" & ae.CreatedAt), "101|2025-01-01")
    End Sub
End Module
