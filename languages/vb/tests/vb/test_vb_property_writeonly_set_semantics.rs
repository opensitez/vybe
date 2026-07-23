use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: WriteOnly Property & Set Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_property_writeonly_password_setter() {
    let src = r#"
Class UserAccount
    Private _passwordHash As String

    Public WriteOnly Property Password As String
        Set(value As String)
            _passwordHash = "HASH_" & value
        End Set
    End Property

    Public Function CheckPassword(input As String) As Boolean
        Return _passwordHash = "HASH_" & input
    End Function
End Class

Module Program
    Sub Main()
        Dim acc As New UserAccount()
        acc.Password = "Secret123"
        Console.WriteLine(acc.CheckPassword("Secret123"))
        Console.WriteLine(acc.CheckPassword("Wrong"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_property_writeonly_write_side_effects() {
    let src = r#"
Class AuditTracker
    Public AuditLog As String = ""

    Public WriteOnly Property LogEntry As String
        Set(value As String)
            AuditLog &= "[" & value & "];"
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim tracker As New AuditTracker()
        tracker.LogEntry = "Event1"
        tracker.LogEntry = "Event2"
        Console.WriteLine(tracker.AuditLog)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[Event1];[Event2];"]);
}

#[test]
fn test_vb_property_set_access_modifier_narrowing() {
    let src = r#"
Class ProtectedSetProp
    Public Property Title As String
        Get
            Return _title
        End Get
        Protected Set(value As String)
            _title = value
        End Set
    End Property

    Private _title As String

    Public Sub New(t As String)
        Title = t
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New ProtectedSetProp("Initial")
        Console.WriteLine(p.Title)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Initial"]);
}
