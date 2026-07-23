use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: StringBuilder Capacity & Advanced Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_sb_capacity_initial_and_expansion() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder(10)
        Console.WriteLine(sb.Capacity >= 10)
        sb.Append("123456789012345")
        Console.WriteLine(sb.Capacity > 10)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_sb_max_capacity_limit() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder(5, 20)
        Console.WriteLine(sb.MaxCapacity)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_sb_append_join_array() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendJoin(", "c, New String() {"Alpha", "Beta", "Gamma"})
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alpha, Beta, Gamma"]);
}

#[test]
fn test_vb_sb_append_format_line() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendFormat("Item {0}: {1:C}", 1, 100)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item 1: $100.00"]);
}

#[test]
fn test_vb_sb_replace_substring() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Hello World")
        sb.Replace("World", "VB.NET")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello VB.NET"]);
}

#[test]
fn test_vb_sb_replace_range_index() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("foo foo foo")
        sb.Replace("foo", "bar", 0, 7)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["bar bar foo"]);
}

#[test]
fn test_vb_sb_insert_at_index() {
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
fn test_vb_sb_remove_range() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Hello Beautiful World")
        sb.Remove(5, 10)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_sb_clear_method() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Some Content")
        sb.Clear()
        Console.WriteLine(sb.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_sb_indexer_get_set() {
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
fn test_vb_sb_copy_to_char_array() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Hello World")
        Dim target(4) As Char
        sb.CopyTo(0, target, 0, 5)
        Console.WriteLine(New String(target))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello"]);
}

#[test]
fn test_vb_sb_equals_string_builder() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb1 As New StringBuilder("Test")
        Dim sb2 As New StringBuilder("Test")
        Console.WriteLine(sb1.Equals(sb2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_sb_ensure_capacity() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        Dim cap As Integer = sb.EnsureCapacity(100)
        Console.WriteLine(cap >= 100)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_sb_append_char_count() {
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
fn test_vb_sb_append_line_empty() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendLine("Line 1")
        sb.AppendLine()
        sb.AppendLine("Line 2")
        Console.WriteLine(sb.ToString().Contains(Environment.NewLine))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_sb_chunk_enumerator() {
    let src = r#"
Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Chunk Enumerator Test")
        Dim count As Integer = 0
        For Each chunk In sb.GetChunks()
            count += 1
        Next
        Console.WriteLine(count > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
