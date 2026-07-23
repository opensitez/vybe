use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Double.TryParse, NumberStyles & CultureInfo Parsing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_double_try_parse_invariant_culture() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        Dim ok = Double.TryParse("1234.56", NumberStyles.Number, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1234.56"]);
}

#[test]
fn test_vb_double_try_parse_german_comma_decimal_separator() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim deCulture As New CultureInfo("de-DE")
        Dim val As Double
        ' German culture uses comma ',' as decimal separator and dot '.' as group separator!
        Dim ok = Double.TryParse("1.234,56", NumberStyles.Number, deCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1234.56"]);
}

#[test]
fn test_vb_double_try_parse_currency_style() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        Dim ok = Double.TryParse("$1,234.50", NumberStyles.Currency, CultureInfo.GetCultureInfo("en-US"), val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1234.5"]);
}

#[test]
fn test_vb_double_try_parse_hex_specifier() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Long
        Dim ok = Long.TryParse("FF00", NumberStyles.HexNumber, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|65280"]);
}

#[test]
fn test_vb_double_try_parse_exponent_notation() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        Dim ok = Double.TryParse("1.5e3", NumberStyles.Float, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1500"]);
}

#[test]
fn test_vb_double_try_parse_parentheses_negative_number() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        ' AllowParentheses parses (100.5) as -100.5!
        Dim ok = Double.TryParse("(100.5)", NumberStyles.AllowParentheses Or NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|-100.5"]);
}

#[test]
fn test_vb_double_try_parse_leading_trailing_whitespace() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        Dim ok = Double.TryParse("   42.75   ", NumberStyles.Float, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|42.75"]);
}

#[test]
fn test_vb_double_try_parse_invalid_string_returns_false() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double = 999.0
        Dim ok = Double.TryParse("NotANumber", val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|0"]);
}

#[test]
fn test_vb_double_try_parse_null_returns_false() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim val As Double
        Dim ok = Double.TryParse(Nothing, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|0"]);
}

#[test]
fn test_vb_double_try_parse_special_nan_infinity_strings() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim valNaN, valInf As Double
        Dim okNaN = Double.TryParse("NaN", NumberStyles.Any, CultureInfo.InvariantCulture, valNaN)
        Dim okInf = Double.TryParse("Infinity", NumberStyles.Any, CultureInfo.InvariantCulture, valInf)
        Console.WriteLine(okNaN & "|" & Double.IsNaN(valNaN) & "|" & okInf & "|" & Double.IsPositiveInfinity(valInf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True|True"]);
}

#[test]
fn test_vb_double_parse_throws_format_exception_on_invalid() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Try
            Double.Parse("InvalidDouble", CultureInfo.InvariantCulture)
        Catch ex As FormatException
            Console.WriteLine("FormatException Caught on Parse Invalid Double")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["FormatException Caught on Parse Invalid Double"]
    );
}

#[test]
fn test_vb_double_parse_throws_overflow_exception_on_too_large() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Try
            Double.Parse("1e309", CultureInfo.InvariantCulture)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException Caught on Double Overflow")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["OverflowException Caught on Double Overflow"]
    );
}

#[test]
fn test_vb_double_try_parse_custom_number_format_info() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim nfi As New NumberFormatInfo() With {
            .NumberDecimalSeparator = "~",
            .NumberGroupSeparator = "'"
        }
        Dim val As Double
        Dim ok = Double.TryParse("1'000~50", NumberStyles.Number, nfi, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1000.5"]);
}

#[test]
fn test_vb_single_try_parse_float_precision() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Single
        Dim ok = Single.TryParse("3.14159", NumberStyles.Float, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|3.14159"]);
}

#[test]
fn test_vb_double_try_parse_french_spaces_group_separator() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim frCulture As New CultureInfo("fr-FR")
        Dim val As Double
        Dim ok = Double.TryParse("1 234,56", NumberStyles.Number, frCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1234.56"]);
}

#[test]
fn test_vb_double_try_parse_trailing_sign() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        Dim ok = Double.TryParse("50.5-", NumberStyles.AllowTrailingSign Or NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|-50.5"]);
}

#[test]
fn test_vb_double_try_parse_span_char_buffer() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim span As ReadOnlySpan(Of Char) = "789.123".ToCharArray()
        Dim val As Double
        Dim ok = Double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|789.123"]);
}

#[test]
fn test_vb_double_try_parse_leading_plus_sign() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        Dim ok = Double.TryParse("+99.9", NumberStyles.AllowLeadingSign Or NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|99.9"]);
}

#[test]
fn test_vb_double_try_parse_strict_no_thousands_separators() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        ' Disallow Thousands separators strictly
        Dim ok = Double.TryParse("1,000.0", NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_double_to_string_culture_formatting_roundtrip() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim orig As Double = 123456.789
        Dim formatted = orig.ToString("N3", CultureInfo.InvariantCulture)
        Dim restored As Double
        Dim ok = Double.TryParse(formatted, NumberStyles.Number, CultureInfo.InvariantCulture, restored)
        Console.WriteLine(formatted & "|" & (orig = restored))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123,456.789|True"]);
}
