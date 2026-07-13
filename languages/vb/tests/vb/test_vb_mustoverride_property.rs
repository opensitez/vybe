use super::helpers::run_vb;

#[test]
fn mustoverride_property() {
    let out = run_vb(
        r#"
MustInherit Class Config
    Public MustOverride Property ConnectionString As String
End Class

Class AppConfig
    Inherits Config
    
    Private _conn As String = "Server=Local;"
    
    Public Overrides Property ConnectionString As String
        Get
            Return _conn
        End Get
        Set(value As String)
            _conn = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As Config = New AppConfig()
        Console.WriteLine(c.ConnectionString)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Server=Local;"]);
}
