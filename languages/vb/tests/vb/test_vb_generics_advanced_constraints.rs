use super::helpers::run_vb;

#[test]
fn generics_multiple_constraints() {
    let out = run_vb(
        r#"
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
        Console.WriteLine(u.Id)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1"]);
}
