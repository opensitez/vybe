use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ReadOnly Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn readonly_property_basic() {
    let out = run_vb(
        r#"
Class Configuration
    Private _maxItems As Integer = 100
    
    Public ReadOnly Property MaxItems As Integer
        Get
            Return _maxItems
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim config As New Configuration()
        Console.WriteLine(config.MaxItems)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn readonly_property_auto_implemented() {
    let out = run_vb(
        r#"
Class User
    Public ReadOnly Property ID As Integer = 12345
End Class

Module M
    Sub Main()
        Dim u As New User()
        Console.WriteLine(u.ID)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn readonly_property_constructor_assignment() {
    let out = run_vb(
        r#"
Class Connection
    Public ReadOnly Property ConnectionString As String
    
    Public Sub New(connStr As String)
        ConnectionString = connStr ' Assignment allowed in constructor for auto-prop
    End Sub
End Class

Module M
    Sub Main()
        Dim conn As New Connection("Server=MyServer")
        Console.WriteLine(conn.ConnectionString)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Server=MyServer"]);
}

#[test]
fn readonly_property_shadowing() {
    let out = run_vb(
        r#"
Class BaseConfig
    Public ReadOnly Property Role As String
        Get
            Return "Guest"
        End Get
    End Property
End Class

Class AdminConfig
    Inherits BaseConfig
    Public Shadows ReadOnly Property Role As String
        Get
            Return "Admin"
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim c As New AdminConfig()
        Console.WriteLine(c.Role)
        
        Dim b As BaseConfig = c
        Console.WriteLine(b.Role)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Admin", "Guest"]);
}
