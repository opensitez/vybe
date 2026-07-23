use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String PadLeft & PadRight Alignment Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_padleft_spaces_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "123"
        Console.WriteLine("'" & s.PadLeft(6) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'   123'"]);
}

#[test]
fn test_vb_string_padleft_custom_char_zeroes() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "42"
        Console.WriteLine(s.PadLeft(5, "0"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["00042"]);
}

#[test]
fn test_vb_string_padleft_custom_char_asterisk() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "99.99"
        Console.WriteLine(s.PadLeft(10, "*"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["*****99.99"]);
}

#[test]
fn test_vb_string_padleft_exact_length_no_change() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "12345"
        Console.WriteLine(s.PadLeft(5, "0"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12345"]);
}

#[test]
fn test_vb_string_padleft_smaller_width_no_change() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "123456789"
        Console.WriteLine(s.PadLeft(5, "0"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123456789"]);
}

#[test]
fn test_vb_string_padleft_empty_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = ""
        Console.WriteLine("'" & s.PadLeft(4, "-"c) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'----'"]);
}

#[test]
fn test_vb_string_padright_spaces_basic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Item"
        Console.WriteLine("'" & s.PadRight(10) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'Item      '"]);
}

#[test]
fn test_vb_string_padright_custom_char_dots() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Chapter 1"
        Console.WriteLine(s.PadRight(20, "."c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Chapter 1..........."]);
}

#[test]
fn test_vb_string_padright_exact_length_no_change() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "FullLength"
        Console.WriteLine(s.PadRight(10, "="c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FullLength"]);
}

#[test]
fn test_vb_string_padright_smaller_width_no_change() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "LongTextHere"
        Console.WriteLine(s.PadRight(4, "="c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LongTextHere"]);
}

#[test]
fn test_vb_string_padright_empty_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = ""
        Console.WriteLine("'" & s.PadRight(4, "*"c) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'****'"]);
}

#[test]
fn test_vb_string_pad_combined_center_alignment() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Title"
        ' Center in width 11: 3 spaces left, 3 spaces right
        Dim centered As String = s.PadLeft(8).PadRight(11)
        Console.WriteLine("'" & centered & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'   Title   '"]);
}

#[test]
fn test_vb_string_padleft_special_characters() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "End"
        Console.WriteLine(s.PadLeft(6, ">"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec![">>>End"]);
}

#[test]
fn test_vb_string_padright_special_characters() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Start"
        Console.WriteLine(s.PadRight(8, "<"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Start<<<"]);
}

#[test]
fn test_vb_string_padleft_number_formatting() {
    let src = r#"
Module Program
    Sub Main()
        Dim i As Integer = 7
        Console.WriteLine(i.ToString().PadLeft(3, "0"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["007"]);
}

#[test]
fn test_vb_string_padright_table_column_formatting() {
    let src = r#"
Module Program
    Sub Main()
        Dim col1 As String = "Name"
        Dim col2 As String = "Age"
        Dim row1Name As String = "Alice"
        Dim row1Age As String = "30"

        Console.WriteLine(col1.PadRight(10) & "|" & col2.PadLeft(5))
        Console.WriteLine(row1Name.PadRight(10) & "|" & row1Age.PadLeft(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Name      |  Age", "Alice     |   30"]);
}

#[test]
fn test_vb_string_padleft_large_width() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "X"
        Console.WriteLine(s.PadLeft(50, "."c).Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_string_padright_null_char_padding() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "AB"
        Dim padded As String = s.PadRight(4, ChrW(0))
        Console.WriteLine(padded.Length)
        Console.WriteLine(AscW(padded(2)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4", "0"]);
}

#[test]
fn test_vb_string_padleft_unicode_padding_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "OK"
        Console.WriteLine(s.PadLeft(5, "★"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["★★★OK"]);
}

#[test]
fn test_vb_string_padright_unicode_padding_char() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As String = "Star"
        Console.WriteLine(s.PadRight(7, "★"c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Star★★★"]);
}
