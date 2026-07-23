use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Text.StringBuilder In-Place Mutations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_builder_append_and_append_line() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.Append("Hello ").AppendLine("World")
        Console.WriteLine(sb.ToString().Contains("Hello World"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_builder_replace_string() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("foo bar foo")
        sb.Replace("foo", "baz")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["baz bar baz"]);
}

#[test]
fn test_vb_string_builder_replace_char() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("a-b-c")
        sb.Replace("-"c, "_"c)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["a_b_c"]);
}

#[test]
fn test_vb_string_builder_insert_at_index() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("AC")
        sb.Insert(1, "B")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ABC"]);
}

#[test]
fn test_vb_string_builder_remove_range() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Hello World")
        sb.Remove(5, 6) ' Remove " World"
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello"]);
}

#[test]
fn test_vb_string_builder_clear_resets_length() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Data")
        sb.Clear()
        Console.WriteLine(sb.Length & "|" & (sb.ToString() = ""))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|True"]);
}

#[test]
fn test_vb_string_builder_capacity_and_max_capacity() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder(10, 100)
        sb.Append("0123456789")
        Console.WriteLine(sb.Capacity & "|" & sb.MaxCapacity)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|100"]);
}

#[test]
fn test_vb_string_builder_exceed_max_capacity_throws() {
    let src = r#"
Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder(5, 5)
        Try
            sb.Append("ExceedMaxCapacity")
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentOutOfRangeException Caught"]);
}

#[test]
fn test_vb_string_builder_indexer_get_set_char() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Cat")
        sb(0) = "B"c
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bat"]);
}

#[test]
fn test_vb_string_builder_append_format() {
    let src = r#"
Imports System.Globalization
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendFormat(CultureInfo.InvariantCulture, "ID: {0}, Val: {1:F2}", 101, 45.678)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ID: 101, Val: 45.68"]);
}

#[test]
fn test_vb_string_builder_append_join() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendJoin(", ", {"A", "B", "C"})
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A, B, C"]);
}

#[test]
fn test_vb_string_builder_copy_to_char_array() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("VisualBasic")
        Dim buffer(5) As Char
        sb.CopyTo(0, buffer, 0, 6)
        Console.WriteLine(New String(buffer))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Visual"]);
}

#[test]
fn test_vb_string_builder_ensure_capacity() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        Dim newCap = sb.EnsureCapacity(500)
        Console.WriteLine(newCap >= 500 & "|" & sb.Capacity >= 500)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_string_builder_append_substring_overload() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.Append("Hello World", 0, 5)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello"]);
}

#[test]
fn test_vb_string_builder_equals_instance_comparison() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb1 As New StringBuilder("Text")
        Dim sb2 As New StringBuilder("Text")
        ' StringBuilder.Equals checks structural capacity/content equality!
        Console.WriteLine(sb1.Equals(sb2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_builder_chained_mutations() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("123")
        sb.Append("456").Insert(0, "0").Remove(2, 2).Replace("5", "X")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["014X6"]);
}

#[test]
fn test_vb_string_builder_append_char_count() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.Append("*"c, 5)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["*****"]);
}

#[test]
fn test_vb_string_builder_set_length_truncate_or_expand() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("ABCDE")
        sb.Length = 3
        Console.WriteLine(sb.ToString())
        sb.Length = 5
        Console.WriteLine(sb.Length & "|" & CInt(sb(3)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ABC", "5|0"]);
}

#[test]
fn test_vb_string_builder_replace_with_start_index_count() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("cat cat cat")
        sb.Replace("cat", "dog", 4, 7) ' Replace starting at index 4 for length 7
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["cat dog dog"]);
}

#[test]
fn test_vb_string_builder_append_object_null() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.Append(CType(Nothing, Object))
        Console.WriteLine(sb.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}
