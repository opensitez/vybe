use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Composite String Formatting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_format_alignment_right() {
    let src = r#"
Module Program
    Sub Main()
        Dim formatted As String = String.Format("{0,10}", "Test")
        Console.WriteLine("'" & formatted & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'      Test'"]);
}

#[test]
fn test_vb_format_alignment_left() {
    let src = r#"
Module Program
    Sub Main()
        Dim formatted As String = String.Format("{0,-10}", "Test")
        Console.WriteLine("'" & formatted & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["'Test      '"]);
}

#[test]
fn test_vb_format_hexadecimal_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 255
        Console.WriteLine(String.Format("{0:X4}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["00FF"]);
}

#[test]
fn test_vb_format_currency_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Decimal = 1234.5D
        Console.WriteLine(String.Format("{0:C2}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["$1,234.50"]);
}

#[test]
fn test_vb_format_fixed_point_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Double = 3.14159
        Console.WriteLine(String.Format("{0:F2}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14"]);
}

#[test]
fn test_vb_format_exponential_scientific_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Double = 12345.678
        Console.WriteLine(String.Format("{0:E2}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.23E+004"]);
}

#[test]
fn test_vb_format_general_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Double = 0.0000123
        Console.WriteLine(String.Format("{0:G}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.23E-05"]);
}

#[test]
fn test_vb_format_number_grouped_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Long = 1000000000L
        Console.WriteLine(String.Format("{0:N0}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,000,000,000"]);
}

#[test]
fn test_vb_format_percent_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Double = 0.75
        Console.WriteLine(String.Format("{0:P0}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["75%"]);
}

#[test]
fn test_vb_format_custom_digit_placeholder() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Double = 12.3
        Console.WriteLine(String.Format("{0:000.00}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["012.30"]);
}

#[test]
fn test_vb_format_custom_zero_placeholder() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Double = 12.345
        Console.WriteLine(String.Format("{0:###.##}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12.35"]);
}

#[test]
fn test_vb_format_section_positive_negative_zero() {
    let src = r#"
Module Program
    Sub Main()
        Dim pos As Double = 5.0
        Dim neg As Double = -5.0
        Dim zero As Double = 0.0
        Dim fmt As String = "{0:Pos:#;Neg:#;Zero}"
        Console.WriteLine(String.Format(fmt, pos))
        Console.WriteLine(String.Format(fmt, neg))
        Console.WriteLine(String.Format(fmt, zero))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Pos:5", "Neg:5", "Zero"]);
}

#[test]
fn test_vb_format_escaped_curly_braces() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 42
        Console.WriteLine(String.Format("{{Value: {0}}}", val))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["{Value: 42}"]);
}

#[test]
fn test_vb_format_multiple_arguments() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(String.Format("{0} + {1} = {2}", 2, 3, 5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2 + 3 = 5"]);
}

#[test]
fn test_vb_format_reordered_arguments() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(String.Format("{2}, {1}, {0}", "A", "B", "C"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["C, B, A"]);
}
