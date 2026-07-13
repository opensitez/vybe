use super::helpers::run_vb;

#[test]
fn extension_methods_inheritance() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices

Interface IEntity
    ReadOnly Property Id As Integer
End Interface

Class User
    Implements IEntity
    Public ReadOnly Property Id As Integer Implements IEntity.Id
        Get
            Return 42
        End Get
    End Property
End Class

Module ExtensionMethods
    <Extension()>
    Public Function GetIdentifier(entity As IEntity) As String
        Return "Entity-" & entity.Id.ToString()
    End Function
End Module

Module M
    Sub Main()
        Dim u As New User()
        ' Extension method on interface
        Console.WriteLine(u.GetIdentifier())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Entity-42"]);
}
