use super::helpers::run_vb;

#[test]
fn interpolation_basic_matrix_simple_expression() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim name As String = "vb"
        Dim value As Integer = 7
        Dim s As String = $"name={name} value={value}"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["name=vb value=7"]);
}

#[test]
fn interpolation_basic_matrix_type_conversion_and_alignment() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Double = 1.2
        Dim s1 As String = $"{x,8:F1}"
        Dim s2 As String = $"{x,8:F1}!"
        Dim n As Integer = 42
        Dim s3 As String = $"[{n,5}]"

        Console.WriteLine(s1.Length)
        Console.WriteLine(s2)
        Console.WriteLine(s3)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8", "   1.2!", "[   42]"]);
}

#[test]
fn interpolation_basic_matrix_nested_expression_with_if() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim score As Integer = 85
        Dim grade As String = $"{If(score >= 90, "A", If(score >= 80, "B", "C"))}"

        Console.WriteLine(grade)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["B"]);
}

#[test]
fn interpolation_basic_matrix_escape_braces() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim name As String = "json"
        Dim s As String = $"{{\"name\": \"{name}\"}}"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["{\"name\": \"json\"}"]);
}

#[test]
fn interpolation_basic_matrix_multiline_interpolation() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As Integer = 1
        Dim b As Integer = 2

        Dim block As String = $"{a}+{b}={a + b}"
        Console.WriteLine(block)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1+2=3"]);
}

#[test]
fn interpolation_basic_matrix_date_formatting() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim dt As Date = #2026-07-21 08:15:00#
        Dim s As String = $"{dt:yyyy-MM-dd}"
        Dim t As String = $"{dt:HH:mm}"

        Console.WriteLine(s)
        Console.WriteLine(t)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2026-07-21", "08:15"]);
}
