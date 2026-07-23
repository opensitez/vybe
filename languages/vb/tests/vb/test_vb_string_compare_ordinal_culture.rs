use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Compare, Ordinal & Culture Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_compare_ordinal_case_sensitive() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As Integer = String.Compare("abc", "ABC", StringComparison.Ordinal)
        Console.WriteLine(res > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_compare_ordinal_ignore_case() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As Integer = String.Compare("abc", "ABC", StringComparison.OrdinalIgnoreCase)
        Console.WriteLine(res = 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_equals_ordinal_ignore_case() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As String = "Hello"
        Dim s2 As String = "HELLO"
        Console.WriteLine(String.Equals(s1, s2, StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_starts_with_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Visual Basic"
        Console.WriteLine(s.StartsWith("visual", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_ends_with_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Visual Basic"
        Console.WriteLine(s.EndsWith("BASIC", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_index_of_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "The quick brown Fox"
        Dim idx As Integer = s.IndexOf("fox", StringComparison.OrdinalIgnoreCase)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_string_last_index_of_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Foo Bar FOO Bar"
        Dim idx As Integer = s.LastIndexOf("foo", StringComparison.OrdinalIgnoreCase)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_string_compare_to_instance() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As String = "apple"
        Dim s2 As String = "banana"
        Console.WriteLine(s1.CompareTo(s2) < 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_compare_ordinal_length_mismatch() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As Integer = String.Compare("abc", "abcd", StringComparison.Ordinal)
        Console.WriteLine(res < 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_contains_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "DotNet Core Framework"
        Console.WriteLine(s.Contains("core", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_invariant_culture_lower_upper() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "i"
        Console.WriteLine(s.ToUpperInvariant())
        Console.WriteLine(s.ToLowerInvariant())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["I", "i"]);
}

#[test]
fn test_vb_string_compare_null_handling() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As String = Nothing
        Dim s2 As String = "Test"
        Console.WriteLine(String.Compare(s1, s2) < 0)
        Console.WriteLine(String.Compare(s1, Nothing) = 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_string_equality_operator_case_sensitivity() {
    let src = r#"
Option Compare Binary

Module Program
    Sub Main()
        Dim s1 As String = "abc"
        Dim s2 As String = "ABC"
        Console.WriteLine(s1 = s2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_string_equality_operator_option_compare_text() {
    let src = r#"
Option Compare Text

Module Program
    Sub Main()
        Dim s1 As String = "abc"
        Dim s2 As String = "ABC"
        Console.WriteLine(s1 = s2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_compare_substring_length() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As String = "Hello World"
        Dim s2 As String = "Hello Universe"
        Dim res As Integer = String.Compare(s1, 0, s2, 0, 5, StringComparison.Ordinal)
        Console.WriteLine(res = 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
