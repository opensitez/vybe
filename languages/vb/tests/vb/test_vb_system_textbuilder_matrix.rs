use super::helpers::run_vb;

#[test]
fn textbuilder_append_and_to_string() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder()
        sb.Append("Hello")
        sb.Append(", ")
        sb.Append("World")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["Hello, World"]);
}

#[test]
fn textbuilder_append_line_adds_line_terminator() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendLine("line1")
        sb.Append("line2")
        Console.WriteLine(sb.ToString().Contains(vbCrLf))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn textbuilder_append_format_is_deterministic() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder()
        sb.AppendFormat("x={0}; y={1}", 1, "two")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["x=1; y=two"]);
}

#[test]
fn textbuilder_insert_and_remove() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("abc")
        sb.Insert(1, "X")
        sb.Remove(2, 1)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["aXc"]);
}

#[test]
fn textbuilder_replace_targets_all_matches() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("banana")
        sb.Replace("a", "o")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["bonono"]);
}

#[test]
fn textbuilder_clear_resets_length_only() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("payload")
        Console.WriteLine(sb.Length)
        sb.Clear()
        Console.WriteLine(sb.Length)
        Console.WriteLine(sb.Capacity >= 7)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["7", "0", "True"]);
}

#[test]
fn textbuilder_copy_to_buffer() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("xyz")
        Dim chars(2) As Char
        sb.CopyTo(0, chars, 0, 3)
        Console.WriteLine(chars(0))
        Console.WriteLine(chars(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["x", "z"]);
}

#[test]
fn textbuilder_chars_read_index() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("delta")
        Console.WriteLine(sb.Chars(1))
        sb.Chars(1) = "E"c
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["e", "dEeta"]);
}

#[test]
fn textbuilder_capacity_growth() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder(2)
        sb.Append("abcdef")
        Console.WriteLine(sb.Length)
        Console.WriteLine(sb.Capacity >= 2)
    End Module
End Module
"#,
    );

    assert_eq!(out, vec!["6", "True"]);
}

#[test]
fn textbuilder_ensure_capacity_returns_new_value() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("abc")
        Dim oldCapacity As Integer = sb.Capacity
        Dim newCapacity As Integer = sb.EnsureCapacity(20)
        Console.WriteLine(newCapacity >= 20)
        Console.WriteLine(sb.Capacity >= 20)
        Console.WriteLine(sb.Capacity >= oldCapacity)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn textbuilder_append_multiple_types() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder()
        sb.Append(True)
        sb.Append("-")
        sb.Append(12.5D)
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True-12.5"]);
}

#[test]
fn textbuilder_replace_range_and_insert() {
    let out = run_vb(
        r#"
Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("abcde")
        sb.Remove(1, 3)
        sb.Insert(1, "XYZ")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["aXYZe"]);
}
