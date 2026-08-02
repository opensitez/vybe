' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_inheritance_specialization
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface IReadRepository(Of T)
    Function GetById(id As Integer) As T
End Interface

Interface IWriteRepository(Of T)
    Sub Save(entity As T)
End Interface

Interface IRepository(Of T)
    Inherits IReadRepository(Of T), IWriteRepository(Of T)
End Interface

Class UserRepo
    Implements IRepository(Of String)
    Private user As String = ""
    Public Function GetById(id As Integer) As String Implements IReadRepository(Of String).GetById
        Return user
    End Function
    Public Sub Save(entity As String) Implements IWriteRepository(Of String).Save
        user = entity
    End Sub
End Class

Module Program
    Sub Main()
        Dim repo As IRepository(Of String) = New UserRepo()
        repo.Save("Alice")
        __Check(CStr(repo.GetById(1)), "Alice")
    End Sub
End Module
