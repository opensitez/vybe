use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ByRef / Out Parameters (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn byref_out_parameter() {
    let out = run_vb(
        r#"
Imports System.Runtime.InteropServices

Module M
    ' VB supports <Out> attribute for interop / pseudo-out parameters
    Sub GetValues(ByRef a As Integer, <Out> ByRef b As Integer)
        a = 10
        b = 20
    End Sub

    Sub Main()
        Dim x, y As Integer
        GetValues(x, y)
        Console.WriteLine(x.ToString() & " " & y.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10 20"]);
}
