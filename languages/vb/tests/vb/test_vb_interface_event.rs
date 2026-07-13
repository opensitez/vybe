use super::helpers::run_vb;

#[test]
fn interface_event() {
    let out = run_vb(
        r#"
Interface INotify
    Event Raised As System.EventHandler
End Interface

Class Notifier
    Implements INotify
    
    Public Event Raised As System.EventHandler Implements INotify.Raised
    
    Public Sub Trigger()
        RaiseEvent Raised(Me, System.EventArgs.Empty)
    End Sub
End Class

Module M
    Sub Main()
        Dim n As New Notifier()
        AddHandler n.Raised, Sub() Console.WriteLine("InterfaceEventRaised")
        n.Trigger()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["InterfaceEventRaised"]);
}
