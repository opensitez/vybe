use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Substring, IndexOf & LastIndexOf Comprehensive Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_substring_basic_start_only() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Hello World"
        Console.WriteLine(s.Substring(6))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["World"]);
}

#[test]
fn test_vb_string_substring_start_and_length() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "VisualBasic.NET"
        Console.WriteLine(s.Substring(0, 6))
        Console.WriteLine(s.Substring(6, 5))
        Console.WriteLine(s.Substring(11, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Visual", "Basic", ".NET"]);
}

#[test]
fn test_vb_string_substring_zero_length() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Test"
        Console.WriteLine("'" & s.Substring(2, 0) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["''"]);
}

#[test]
fn test_vb_string_substring_full_length() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Complete"
        Console.WriteLine(s.Substring(0, s.Length))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Complete"]);
}

#[test]
fn test_vb_string_indexof_char_single() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "banana"
        Console.WriteLine(s.IndexOf("a"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_string_indexof_char_start_index() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "banana"
        Console.WriteLine(s.IndexOf("a"c, 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_string_indexof_char_start_and_count() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "banana"
        Console.WriteLine(s.IndexOf("a"c, 2, 2))
        Console.WriteLine(s.IndexOf("a"c, 2, 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "-1"]);
}

#[test]
fn test_vb_string_indexof_string_ordinal() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Quick Brown Fox"
        Console.WriteLine(s.IndexOf("Brown"))
        Console.WriteLine(s.IndexOf("brown"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6", "-1"]);
}

#[test]
fn test_vb_string_indexof_string_comparison_enum() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "Quick Brown Fox"
        Console.WriteLine(s.IndexOf("brown", StringComparison.OrdinalIgnoreCase))
        Console.WriteLine(s.IndexOf("BROWN", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6", "6"]);
}

#[test]
fn test_vb_string_indexof_string_start_and_comparison() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "Test Test Test"
        Console.WriteLine(s.IndexOf("test", 5, StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_string_lastindexof_char_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "banana"
        Console.WriteLine(s.LastIndexOf("a"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_string_lastindexof_char_start_index() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "banana"
        Console.WriteLine(s.LastIndexOf("a"c, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_string_lastindexof_char_start_and_count() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "banana"
        Console.WriteLine(s.LastIndexOf("a"c, 4, 3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_string_lastindexof_string_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "foo bar foo baz foo"
        Console.WriteLine(s.LastIndexOf("foo"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_string_lastindexof_string_comparison_ignore_case() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "FOO bar Foo baz FOO"
        Console.WriteLine(s.LastIndexOf("foo", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_string_indexofany_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "hello world"
        Dim targets As Char() = {"o"c, "w"c}
        Console.WriteLine(s.IndexOfAny(targets))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_string_indexofany_start_index() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "hello world"
        Dim targets As Char() = {"o"c, "w"c}
        Console.WriteLine(s.IndexOfAny(targets, 5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_string_lastindexofany_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "hello world"
        Dim targets As Char() = {"l"c, "e"c}
        Console.WriteLine(s.LastIndexOfAny(targets))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9"]);
}

#[test]
fn test_vb_string_contains_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "abcdef"
        Console.WriteLine(s.Contains("c"c))
        Console.WriteLine(s.Contains("z"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_string_contains_string_case_sensitive() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "DotNet Framework"
        Console.WriteLine(s.Contains("Net"))
        Console.WriteLine(s.Contains("net"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_string_contains_string_comparison() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "DotNet Framework"
        Console.WriteLine(s.Contains("net", StringComparison.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_startswith_endswith_overloads() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim s As String = "https://example.com/index.html"
        Console.WriteLine(s.StartsWith("HTTPS", StringComparison.OrdinalIgnoreCase))
        Console.WriteLine(s.EndsWith(".HTML", StringComparison.OrdinalIgnoreCase))
        Console.WriteLine(s.StartsWith("http://"))
        Console.WriteLine(s.EndsWith(".php"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True", "False", "False"]);
}
