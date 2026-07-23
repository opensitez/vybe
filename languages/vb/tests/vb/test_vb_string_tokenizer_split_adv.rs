use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Advanced String Tokenization & Splitting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_split_multiple_char_separators() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = "apple,banana;orange:grape"
        Dim parts As String() = text.Split(New Char() {","c, ";"c, ":"c})
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4", "banana"]);
}

#[test]
fn test_vb_string_split_remove_empty_entries() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = "one,,two,,,three"
        Dim parts As String() = text.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "three"]);
}

#[test]
fn test_vb_string_split_trim_entries() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = "  alpha ,  beta  , gamma "
        Dim parts As String() = text.Split(New Char() {","c}, StringSplitOptions.TrimEntries)
        Console.WriteLine("'" & parts(0) & "'")
        Console.WriteLine("'" & parts(1) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'alpha'", "'beta'"]);
}

#[test]
fn test_vb_string_split_string_separator_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = "Foo<BR>Bar<BR>Baz"
        Dim parts As String() = text.Split(New String() {"<BR>"}, StringSplitOptions.None)
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "Bar"]);
}

#[test]
fn test_vb_string_split_count_limit() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = "a,b,c,d,e"
        Dim parts As String() = text.Split(New Char() {","c}, 3)
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "c,d,e"]);
}

#[test]
fn test_vb_string_join_enumerable() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"Red", "Green", "Blue"}
        Dim joined As String = String.Join(" - ", list)
        Console.WriteLine(joined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Red - Green - Blue"]);
}

#[test]
fn test_vb_string_join_array_range() {
    let src = r#"
Module Program
    Sub Main()
        Dim items As String() = {"A", "B", "C", "D", "E"}
        Dim joined As String = String.Join("|", items, 1, 3)
        Console.WriteLine(joined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B|C|D"]);
}

#[test]
fn test_vb_string_concat_objects() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As String = String.Concat("Value: ", 42, " Result: ", True)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Value: 42 Result: True"]);
}

#[test]
fn test_vb_string_concat_enumerable() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim nums As New List(Of Integer) From {1, 2, 3, 4}
        Dim res As String = String.Concat(nums)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1234"]);
}

#[test]
fn test_vb_string_tokenizer_span_lines() {
    let src = r#"
Module Program
    Sub Main()
        Dim multiline As String = "Line1" & vbCrLf & "Line2" & vbLf & "Line3"
        Dim lines As String() = multiline.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
        Console.WriteLine(lines.Length)
        Console.WriteLine(lines(0))
        Console.WriteLine(lines(1))
        Console.WriteLine(lines(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "Line1", "Line2", "Line3"]);
}

#[test]
fn test_vb_string_split_char_implicit() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = "x y z"
        Dim parts As String() = text.Split(" "c)
        Console.WriteLine(parts.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_string_split_trim_and_remove_empty() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = " a ,   , b "
        Dim opts As StringSplitOptions = StringSplitOptions.TrimEntries Or StringSplitOptions.RemoveEmptyEntries
        Dim parts As String() = text.Split(New Char() {","c}, opts)
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(0))
        Console.WriteLine(parts(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "a", "b"]);
}

#[test]
fn test_vb_string_join_generic_objects() {
    let src = r#"
Module Program
    Sub Main()
        Dim dates As DateTime() = {New DateTime(2026, 1, 1), New DateTime(2026, 12, 31)}
        Dim res As String = String.Join(" to ", dates)
        Console.WriteLine(res.Contains("2026"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_string_concat_char_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim chars As Char() = {"H"c, "e"c, "l"c, "l"c, "o"c}
        Dim s As String = New String(chars)
        Console.WriteLine(s)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello"]);
}

#[test]
fn test_vb_string_split_no_match() {
    let src = r#"
Module Program
    Sub Main()
        Dim text As String = "NoSeparators"
        Dim parts As String() = text.Split(","c)
        Console.WriteLine(parts.Length)
        Console.WriteLine(parts(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "NoSeparators"]);
}
