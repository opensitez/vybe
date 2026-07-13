use super::helpers::run_vb;

#[test]
fn readonly_fields_in_constructors() {
    let out = run_vb(
        r#"
Class Config
    Public ReadOnly BaseUrl As String
    
    Public Sub New(url As String)
        BaseUrl = url
    End Sub
    
    ' VB.NET allows assigning to ReadOnly fields in any constructor
    Public Sub New()
        Me.New("http://localhost")
    End Sub
End Class

Module M
    Sub Main()
        Dim c1 As New Config()
        Console.WriteLine(c1.BaseUrl)
        
        Dim c2 As New Config("http://test")
        Console.WriteLine(c2.BaseUrl)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["http://localhost", "http://test"]);
}

#[test]
fn readonly_properties_init() {
    let out = run_vb(
        r#"
Class User
    ' ReadOnly auto-property can be initialized at declaration
    Public ReadOnly Property Id As Integer = 100
    Public ReadOnly Property Name As String
    
    Public Sub New(name As String)
        ' Can also be initialized in constructor
        Me.Name = name
    End Sub
End Class

Module M
    Sub Main()
        Dim u As New User("Alice")
        Console.WriteLine(u.Id)
        Console.WriteLine(u.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100", "Alice"]);
}
