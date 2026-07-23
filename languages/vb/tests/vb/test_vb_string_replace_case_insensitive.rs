use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String.Replace Overloads & Surface Area
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_replace_char_to_char_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "123-456-7890"
        Console.WriteLine(s.Replace("-"c, "."c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123.456.7890"]);
}

#[test]
fn test_vb_string_replace_char_not_found() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Hello World"
        Console.WriteLine(s.Replace("x"c, "y"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_string_replace_string_to_string_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "The quick brown fox jumps over the lazy dog"
        Console.WriteLine(s.Replace("fox", "cat"))
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["The quick brown cat jumps over the lazy dog"]
    );
}

#[test]
fn test_vb_string_replace_string_empty_replacement() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "A B C D E"
        Console.WriteLine(s.Replace(" ", ""))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ABCDE"]);
}

#[test]
fn test_vb_string_replace_string_multiple_occurrences() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "foo bar foo baz foo"
        Console.WriteLine(s.Replace("foo", "qux"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["qux bar qux baz qux"]);
}

#[test]
fn test_vb_string_replace_string_comparison_ordinal() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "Case CASE case"
        Console.WriteLine(s.Replace("CASE", "LOWER", StringComparison.Ordinal))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Case LOWER case"]);
}

#[test]
fn test_vb_string_replace_string_comparison_ordinal_ignore_case() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "Case CASE case"
        Console.WriteLine(s.Replace("case", "MATCH", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["MATCH MATCH MATCH"]);
}

#[test]
fn test_vb_string_replace_overlapping_pattern() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "aaaa"
        Console.WriteLine(s.Replace("aa", "b"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["bb"]);
}

#[test]
fn test_vb_string_replace_with_longer_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "1 2 3"
        Console.WriteLine(s.Replace(" ", " AND "))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1 AND 2 AND 3"]);
}

#[test]
fn test_vb_string_replace_with_shorter_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "apple--banana--cherry"
        Console.WriteLine(s.Replace("--", "-"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["apple-banana-cherry"]);
}

#[test]
fn test_vb_string_replace_single_character_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "hello"
        Console.WriteLine(s.Replace("l", "w"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["hewwo"]);
}

#[test]
fn test_vb_string_replace_newlines() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Line1" & vbCrLf & "Line2"
        Console.WriteLine(s.Replace(vbCrLf, ";"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Line1;Line2"]);
}

#[test]
fn test_vb_string_replace_tabs() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Col1" & vbTab & "Col2"
        Console.WriteLine(s.Replace(vbTab, "    "))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Col1    Col2"]);
}

#[test]
fn test_vb_string_replace_chained_calls() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "a-b-c-d"
        Dim result = s.Replace("a", "1").Replace("b", "2").Replace("c", "3").Replace("d", "4")
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1-2-3-4"]);
}

#[test]
fn test_vb_string_replace_special_regex_chars_literal() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "price: $10.00 (USD)"
        Console.WriteLine(s.Replace("$", "€").Replace("(", "[").Replace(")", "]"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["price: €10.00 [USD]"]);
}

#[test]
fn test_vb_string_replace_unicode_characters() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "α β γ α"
        Console.WriteLine(s.Replace("α", "omega"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["omega β γ omega"]);
}

#[test]
fn test_vb_string_replace_quotes() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Say ""Hello"""
        Console.WriteLine(s.Replace(""""c, "'"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Say 'Hello'"]);
}

#[test]
fn test_vb_string_replace_same_old_and_new_value() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Unchanged"
        Console.WriteLine(s.Replace("Unchanged", "Unchanged"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Unchanged"]);
}

#[test]
fn test_vb_string_replace_entire_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "EntireText"
        Console.WriteLine(s.Replace("EntireText", "ReplacedText"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ReplacedText"]);
}

#[test]
fn test_vb_string_replace_case_insensitive_culture_invariant() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "Straße STRASSE straße"
        Console.WriteLine(s.Replace("STRASSE", "STREET", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Straße STREET straße"]);
}
