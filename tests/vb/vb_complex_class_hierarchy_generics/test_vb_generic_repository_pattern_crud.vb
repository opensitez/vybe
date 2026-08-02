' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_repository_pattern_crud
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

Imports System.Collections.Generic

Interface IEntity
    Property Id As Integer
End Interface

Class User
    Implements IEntity
    Public Property Id As Integer Implements IEntity.Id
    Public Property Name As String
End Class

Class Repository(Of T As {Class, IEntity, New})
    Private items As New List(Of T)()

    Public Sub Add(item As T)
        items.Add(item)
    End Sub

    Public Function GetById(id As Integer) As T
        For Each item In items
            If item.Id = id Then Return item
        Next
        Return Nothing
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As New Repository(Of User)()
        repo.Add(New User With {.Id = 1, .Name = "Alice"})
        repo.Add(New User With {.Id = 2, .Name = "Bob"})

        Dim u = repo.GetById(2)
        Console.WriteLine(u.Name)
    End Sub
End Module
