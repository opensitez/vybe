use super::helpers::run_vb;

#[test]
fn handles_mybase() {
    let out = run_vb(
        r#"
Class Base
    Public Event Processed As EventHandler
    
    Protected Sub Trigger()
        RaiseEvent Processed(Me, EventArgs.Empty)
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' Handles MyBase.Event
    Private Sub OnProcessed(sender As Object, e As EventArgs) Handles MyBase.Processed
        Console.WriteLine("Handled in Derived")
    End Sub
    
    Public Sub DoWork()
        Trigger()
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.DoWork()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Handled in Derived"]);
}
