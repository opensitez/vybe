use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Option Strict Off Late Binding
// ═══════════════════════════════════════════════════════════

#[test]
fn option_strict_off() {
    let out = run_vb(
        r#"
Option Strict Off

Module M
    Sub Main()
        ' With Option Strict Off, implicit narrowing conversions are allowed
        Dim x As Double = 42.5
        Dim y As Integer = x ' Implicit conversion to Integer, rounds to 42
        Console.WriteLine(y)
        
        ' And Late Binding is allowed
        Dim obj As Object = "Hello"
        Console.WriteLine(obj.ToUpper())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "HELLO"]);
}
