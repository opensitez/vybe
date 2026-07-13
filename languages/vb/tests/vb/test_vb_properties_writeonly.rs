use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: WriteOnly Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn writeonly_property_basic() {
    let out = run_vb(
        r#"
Class Logger
    Private _lastMessage As String
    
    Public WriteOnly Property Message As String
        Set(value As String)
            _lastMessage = value
            Console.WriteLine("Logged: " & _lastMessage)
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim log As New Logger()
        log.Message = "System started"
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Logged: System started"]);
}

#[test]
fn writeonly_property_side_effects() {
    let out = run_vb(
        r#"
Class Counter
    Public Total As Integer = 0
    
    Public WriteOnly Property AddAmount As Integer
        Set(value As Integer)
            Total = Total + value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        c.AddAmount = 5
        c.AddAmount = 10
        Console.WriteLine(c.Total)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn writeonly_property_in_interface() {
    let out = run_vb(
        r#"
Interface IPasswordSettable
    WriteOnly Property Password As String
End Interface

Class User
    Implements IPasswordSettable
    
    Public WriteOnly Property Password As String Implements IPasswordSettable.Password
        Set(value As String)
            Console.WriteLine("Password set to length: " & value.Length.ToString())
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim u As IPasswordSettable = New User()
        u.Password = "secret123"
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Password set to length: 9"]);
}
