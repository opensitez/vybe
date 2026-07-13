use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Strings Advanced (Left, Right, Space, String)
// ═══════════════════════════════════════════════════════════

#[test]
fn string_left_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "Programming"
        Console.WriteLine(Left(text, 4))
        Console.WriteLine(Left(text, 100))
        Console.WriteLine(Left(text, 0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Prog", "Programming", ""]);
}

#[test]
fn string_right_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "Programming"
        Console.WriteLine(Right(text, 3))
        Console.WriteLine(Right(text, 100))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["ing", "Programming"]);
}

#[test]
fn string_space_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "A" & Space(3) & "B"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A   B"]);
}

#[test]
fn string_string_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Creates a string of a repeated character
        Console.WriteLine(String(5, "x"c))
        ' Also works with char code
        Console.WriteLine(String(3, 42)) ' 42 is '*'
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["xxxxx", "***"]);
}
