use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interfaces (Explicit Implementation)
// ═══════════════════════════════════════════════════════════

#[test]
fn interface_explicit_implementation_different_name() {
    let out = run_vb(
        r#"
Interface IShape
    Sub Draw()
End Interface

Class Circle
    Implements IShape
    
    ' Explicitly implementing with a different method name
    Private Sub Render() Implements IShape.Draw
        Console.WriteLine("Drawing Circle")
    End Sub
    
    Public Sub Draw()
        Console.WriteLine("This is class Draw, not interface Draw")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Circle()
        c.Draw() ' Calls class method
        
        Dim s As IShape = c
        s.Draw() ' Calls interface method (Render)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["This is class Draw, not interface Draw", "Drawing Circle"]
    );
}

#[test]
fn interface_explicit_property_different_name() {
    let out = run_vb(
        r#"
Interface ICounter
    ReadOnly Property Value As Integer
End Interface

Class MyCounter
    Implements ICounter
    
    Private ReadOnly Property ICounter_Value As Integer Implements ICounter.Value
        Get
            Return 42
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim c As ICounter = New MyCounter()
        Console.WriteLine(c.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
