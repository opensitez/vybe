use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ByRef with Object Late Binding
// ═══════════════════════════════════════════════════════════

#[test]
fn byref_late_binding() {
    let out = run_vb(
        r#"
Option Strict Off

Module M
    Sub ModifyRef(ByRef val As Integer)
        val += 10
    End Sub

    Sub Main()
        Dim obj As Object = 5
        ' Late binding ByRef passes the object, which is unboxed, modified, and re-boxed
        ' Wait, if we pass an Object to an Integer ByRef parameter with Option Strict Off,
        ' VB.NET generates a temporary variable and copies it back.
        ModifyRef(obj)
        Console.WriteLine(obj)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}
