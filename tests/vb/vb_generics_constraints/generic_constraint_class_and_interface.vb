' vybe-test: vb/vb_generics_constraints/generic_constraint_class_and_interface
' origin: languages/vb/tests/vb/test_vb_generics_constraints.rs

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

Interface IIdentifiable
    Function GetID() As Integer
End Interface

' T must be a reference type and implement IIdentifiable
Class Repository(Of T As {Class, IIdentifiable})
    Public Sub Process(item As T)
        __Check(CStr("Processing ID: " & item.GetID().ToString()), "Processing ID: 999")
    End Sub
End Class

Class User
    Implements IIdentifiable
    Public Function GetID() As Integer Implements IIdentifiable.GetID
        Return 999
    End Function
End Class

Module M
    Sub Main()
        Dim r As New Repository(Of User)()
        Dim u As New User()
        r.Process(u)
    End Sub
End Module
