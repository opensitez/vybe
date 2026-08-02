' vybe-test: vb/vb_abstract_class_inheritance_chain/test_vb_mustinherit_constructor_invocation
' origin: languages/vb/tests/vb/test_vb_abstract_class_inheritance_chain.rs

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

MustInherit Class Entity
    Public Property Id As Integer
    Protected Sub New(id As Integer)
        Me.Id = id
    End Sub
End Class

Class User
    Inherits Entity
    Public Property Name As String
    Public Sub New(id As Integer, name As String)
        MyBase.New(id)
        Me.Name = name
    End Sub
End Class

Module Program
    Sub Main()
        Dim u As New User(42, "Alice")
        __Check(CStr(u.Id & ":" & u.Name), "42:Alice")
    End Sub
End Module
