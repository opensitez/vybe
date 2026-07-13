use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generics (Methods)
// ═══════════════════════════════════════════════════════════

#[test]
fn generic_method_basic() {
    let out = run_vb(
        r#"
Module M
    Function Identity(Of T)(value As T) As T
        Return value
    End Function

    Sub Main()
        Console.WriteLine(Identity(Of String)("Test"))
        Console.WriteLine(Identity(Of Integer)(123))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Test", "123"]);
}

#[test]
fn generic_method_type_inference() {
    let out = run_vb(
        r#"
Module M
    Sub Swap(Of T)(ByRef a As T, ByRef b As T)
        Dim temp As T = a
        a = b
        b = temp
    End Sub

    Sub Main()
        Dim x As Integer = 1
        Dim y As Integer = 2
        ' Type parameter omitted, compiler infers (Of Integer)
        Swap(x, y)
        Console.WriteLine(x)
        Console.WriteLine(y)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}
