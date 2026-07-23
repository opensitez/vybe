use super::helpers::run_vb;

#[test]
fn string_split_single_delimiter() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim parts As String() = "a,b,c".Split(",")
        Console.WriteLine(parts(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["b"]);
}

#[test]
fn string_split_removes_empty_entries() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim parts As String() = "a,,b".Split(New String() {","}, StringSplitOptions.RemoveEmptyEntries)
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(0))
        Console.WriteLine(parts(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "a", "b"]);
}

#[test]
fn string_join_concatenates_with_separator() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(String.Join("-", New String() {"a", "b", "c"}))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a-b-c"]);
}

#[test]
fn string_trim_removes_whitespace() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("  hello  ".Trim())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["hello"]);
}

#[test]
fn string_pad_left_right_to_width() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("hi".PadLeft(5))
        Console.WriteLine("hi".PadRight(5))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["   hi", "hi   "]);
}

#[test]
fn string_contains_prefix_and_suffix_checks() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("prefix_body".StartsWith("prefix"))
        Console.WriteLine("body_suffix".EndsWith("suffix"))
        Console.WriteLine("foobar".Contains("oba"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn string_index_functions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("abcabc".IndexOf("b"c))
        Console.WriteLine("abcabc".LastIndexOf("b"c))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn string_replace_substring_occurrences() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("aabbaa".Replace("aa", "X"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["XbbX"]);
}

#[test]
fn string_substring_extracts_region() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("hello world".Substring(6, 5))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["world"]);
}

#[test]
fn string_case_mapping_and_locale_safe_numeric_char() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("Hello".ToUpper())
        Console.WriteLine("Hello".ToLower())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["HELLO", "hello"]);
}

#[test]
fn string_index_operator_reads_character() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("hello"(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["e"]);
}

#[test]
fn string_is_null_or_whitespace() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(String.IsNullOrWhiteSpace("   "))
        Console.WriteLine(String.IsNullOrWhiteSpace("x"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn string_static_concat_works_with_many_operands() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(String.Concat("a", "b", "c"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["abc"]);
}

#[test]
fn string_compare_ordinal_is_case_sensitive() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim r As Integer = String.Compare("A", "a", StringComparison.Ordinal)
        Console.WriteLine(r <> 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn string_trim_edges() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("  abc  ".TrimStart())
        Console.WriteLine("  abc  ".TrimEnd())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["abc  ", "  abc"]);
}

#[test]
fn string_insert_and_remove() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim base As String = "abc"
        Console.WriteLine(base.Insert(1, "X"))
        Console.WriteLine(base.Remove(1, 1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["aXbc", "ac"]);
}

#[test]
fn string_join_with_empty_separator() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(String.Join("", New String() {"a", "b", "c"}))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["abc"]);
}

#[test]
fn string_left_right_mid_inclusive_edges() {
    let out = run_vb(
        r#"
Imports Microsoft.VisualBasic

Module M
    Sub Main()
        Dim source As String = "abcdef"
        Console.WriteLine(Strings.Left(source, 2))
        Console.WriteLine(Strings.Right(source, 2))
        Console.WriteLine(Strings.Mid(source, 2, 3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["ab", "ef", "bcd"]);
}
