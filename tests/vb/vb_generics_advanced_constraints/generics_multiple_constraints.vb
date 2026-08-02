' vybe-test: vb/vb_generics_advanced_constraints/generics_multiple_constraints
' origin: languages/vb/tests/vb/test_vb_generics_advanced_constraints.rs

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
    Property Id As Integer
End Interface

' Multiple constraints: must be a class, have a parameterless constructor, and implement IIdentifiable
Class Repository(Of T As {Class, IIdentifiable, New})
    Public Function CreateNew() As T
        Dim obj As New T()
        obj.Id = 1
        Return obj
    End Function
End Class

Class User
    Implements IIdentifiable
    Public Property Id As Integer Implements IIdentifiable.Id
    Public Property Name As String
End Class

Module M
    Sub Main()
        Dim repo As New Repository(Of User)()
        Dim u = repo.CreateNew()
        __Check(CStr(u.Id), "1")
    End Sub
End Module
