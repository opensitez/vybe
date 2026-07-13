use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: AddressOf with Overloaded Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn addressof_overloads() {
    let out = run_vb(
        r#"
Delegate Sub PrintString(s As String)
Delegate Sub PrintInteger(i As Integer)

Module M
    Sub Print(s As String)
        Console.WriteLine("String: " & s)
    End Sub

    Sub Print(i As Integer)
        Console.WriteLine("Integer: " & i.ToString())
    End Sub

    Sub Main()
        ' AddressOf automatically selects the correct overload based on the target delegate type
        Dim ds As PrintString = AddressOf Print
        Dim di As PrintInteger = AddressOf Print
        
        ds("Hello")
        di(42)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["String: Hello", "Integer: 42"]);
}
