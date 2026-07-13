use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interface Implementation Rules (Implements Clause)
// ═══════════════════════════════════════════════════════════

#[test]
fn interface_implementation_explicit_names() {
    let out = run_vb(
        r#"
Interface IPrinter
    Sub Print()
End Interface

Class ConsolePrinter
    Implements IPrinter
    
    ' In VB.NET, the implementing method name doesn't have to match the interface method name
    ' The Implements clause defines what it implements
    Public Sub Output() Implements IPrinter.Print
        Console.WriteLine("Printed explicitly")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As IPrinter = New ConsolePrinter()
        p.Print()
        
        Dim cp As New ConsolePrinter()
        cp.Output()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Printed explicitly", "Printed explicitly"]);
}
