use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Composite Formatting, Alignment & Format Specifiers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_format_alignment_right_padded() {
    let src = r#"
Module Program
    Sub Main()
        ' Right-aligned in 10-character field
        Dim res = String.Format("[{0,10}]", "Test")
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[      Test]"]);
}

#[test]
fn test_vb_string_format_alignment_left_padded() {
    let src = r#"
Module Program
    Sub Main()
        ' Left-aligned in 10-character field
        Dim res = String.Format("[{0,-10}]", "Test")
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[Test      ]"]);
}

#[test]
fn test_vb_string_format_hexadecimal_specifier() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = String.Format("{0:X4}", 255)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["00FF"]);
}

#[test]
fn test_vb_string_format_currency_specifier() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:C2}", 1234.5)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["¤1,234.50"]);
}

#[test]
fn test_vb_string_format_exponential_specifier() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:E2}", 12345.678)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.23E+004"]);
}

#[test]
fn test_vb_string_format_fixed_point_specifier() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:F3}", 3.14)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.140"]);
}

#[test]
fn test_vb_string_format_general_specifier() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:G}", 123.456)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123.456"]);
}

#[test]
fn test_vb_string_format_number_with_group_separators() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:N2}", 1000000)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,000,000.00"]);
}

#[test]
fn test_vb_string_format_percent_specifier() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:P1}", 0.755)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["75.5%"]);
}

#[test]
fn test_vb_string_format_decimal_leading_zeros() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = String.Format("{0:D5}", 42)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["00042"]);
}

#[test]
fn test_vb_string_format_custom_hash_zero_specifiers() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:###,##0.00}", 1234.5)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,234.50"]);
}

#[test]
fn test_vb_string_format_conditional_three_section_specifier() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim fmt = "{0:+#,##0;-#,##0;ZERO}"
        Console.WriteLine(String.Format(CultureInfo.InvariantCulture, fmt, 50) & "|" & String.Format(CultureInfo.InvariantCulture, fmt, -50) & "|" & String.Format(CultureInfo.InvariantCulture, fmt, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["+50|-50|ZERO"]);
}

#[test]
fn test_vb_string_format_escaped_braces() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = String.Format("{{Value: {0}}}", 99)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["{Value: 99}"]);
}

#[test]
fn test_vb_string_format_multiple_positional_arguments() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = String.Format("{2} - {1} - {0}", "Third", "Second", "First")
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First - Second - Third"]);
}

#[test]
fn test_vb_string_format_alignment_and_format_combined() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim res = String.Format(CultureInfo.InvariantCulture, "[{0,-10:C2}]", 50.0)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[¤50.00    ]"]);
}

#[test]
fn test_vb_string_format_date_time_format_specifiers() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 4, 1, 14, 5, 9)
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd HH:mm:ss}", dt)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-04-01 14:05:09"]);
}

#[test]
fn test_vb_string_format_enum_formatting_g_f_d_x() {
    let src = r#"
Enum Level
    Low = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim l = Level.High
        Console.WriteLine(String.Format("{0:G}|{0:D}|{0:X}", l))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["High|2|00000002"]);
}

#[test]
fn test_vb_string_format_throws_format_exception_missing_arg() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            String.Format("{0} {1}", "OnlyOne")
        Catch ex As FormatException
            Console.WriteLine("FormatException Missing Argument Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FormatException Missing Argument Caught"]);
}

#[test]
fn test_vb_string_format_null_argument_evaluated_as_empty() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = String.Format("Val: [{0}]", CType(Nothing, String))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Val: []"]);
}

#[test]
fn test_vb_string_format_param_array_overload() {
    let src = r#"
Module Program
    Sub Main()
        Dim args As Object() = {"A", "B", "C", "D"}
        Dim res = String.Format("{0}-{1}-{2}-{3}", args)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A-B-C-D"]);
}
