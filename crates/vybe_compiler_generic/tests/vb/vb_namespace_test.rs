use super::helpers::run_vb;

// ============================================================
// Namespace object access (struct_get chains)
// ============================================================

#[test]
fn math_namespace_object() {
    // Math.Floor accessed via namespace object (struct_get on global "math")
    // The compiler currently uses call_import for Math.Floor,
    // but the namespace object also works if accessed dynamically
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Math.Floor(3.7))
        Console.WriteLine(Math.Abs(-5))
        Console.WriteLine(Math.Sqrt(9))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3", "5", "3"]);
}

#[test]
fn console_namespace_object() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine("hello from namespace")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["hello from namespace"]);
}

#[test]
fn convert_namespace_object() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Convert.ToString(42))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn math_pi_constant() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Math.Floor(Math.PI))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3"]);
}
