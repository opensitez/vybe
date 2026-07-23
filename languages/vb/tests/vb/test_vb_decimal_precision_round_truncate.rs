use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Decimal High Precision Arithmetic & Rounding Modes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_decimal_literal_and_precision() {
    let src = r#"
Module Program
    Sub Main()
        Dim dec As Decimal = 123456789.987654321D
        Console.WriteLine(dec.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123456789.987654321"]);
}

#[test]
fn test_vb_decimal_addition_exact_no_floating_point_error() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Decimal = 0.1D
        Dim b As Decimal = 0.2D
        Dim sum = a + b
        Console.WriteLine(sum = 0.3D)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_decimal_division_precision_limit() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Decimal = 1D
        Dim b As Decimal = 3D
        Dim res = a / b
        Console.WriteLine(res.ToString().StartsWith("0.3333333333333333333333333333"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_decimal_round_bankers_rounding() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Decimal.Round with default Banker's Rounding
        Console.WriteLine(Decimal.Round(2.5D) & "|" & Decimal.Round(3.5D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|4"]);
}

#[test]
fn test_vb_decimal_round_midpoint_away_from_zero() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Decimal.Round(2.5D, MidpointRounding.AwayFromZero) & "|" & Decimal.Round(3.5D, MidpointRounding.AwayFromZero))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|4"]);
}

#[test]
fn test_vb_decimal_truncate_drops_fractional_part() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Decimal.Truncate(123.456D) & "|" & Decimal.Truncate(-123.456D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123|-123"]);
}

#[test]
fn test_vb_decimal_ceiling_and_floor() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Decimal.Ceiling(10.1D) & "|" & Decimal.Floor(10.9D) & "|" & Decimal.Floor(-10.1D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11|10|-11"]);
}

#[test]
fn test_vb_decimal_min_max_values() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Decimal.MinValue < 0D & "|" & Decimal.MaxValue > 0D)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_decimal_overflow_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim max = Decimal.MaxValue
        Try
            Dim overflow = max + 1D
        Catch ex As OverflowException
            Console.WriteLine("Decimal OverflowException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Decimal OverflowException Caught"]);
}

#[test]
fn test_vb_decimal_get_bits_roundtrip() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim original As Decimal = 123.45D
        Dim bits = Decimal.GetBits(original)
        Dim restored As New Decimal(bits)
        Console.WriteLine(original = restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_decimal_remainder_mod_operator() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Decimal = 10.5D
        Dim b As Decimal = 3D
        Dim remVal = a Mod b
        Console.WriteLine(remVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.5"]);
}

#[test]
fn test_vb_decimal_comparison_operators() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 10.0D
        Dim d2 As Decimal = 10.00D
        ' Decimal maintains scale but equality operator treats 10.0 and 10.00 as equal!
        Console.WriteLine((d1 = d2) & "|" & (d1 <= d2) & "|" & (d1 >= d2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_decimal_currency_formatting() {
    let src = r#"
Imports System.Globalization

Module Program
    Sub Main()
        Dim price As Decimal = 99.99D
        Console.WriteLine(price.ToString("C2", CultureInfo.InvariantCulture))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["¤99.99"]);
}

#[test]
fn test_vb_decimal_parse_and_try_parse() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Decimal
        Dim ok = Decimal.TryParse("1234.56", NumberStyles.Number, CultureInfo.InvariantCulture, val)
        Console.WriteLine(ok & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1234.56"]);
}

#[test]
fn test_vb_decimal_to_int32_casting_truncates() {
    let src = r#"
Module Program
    Sub Main()
        Dim dec As Decimal = 42.99D
        Dim intVal As Integer = CInt(dec) ' CInt uses banker's rounding for Decimal -> Integer!
        Console.WriteLine(intVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["43"]);
}

#[test]
fn test_vb_decimal_fix_function_behaviour() {
    let src = r#"
Module Program
    Sub Main()
        ' Fix function truncates towards zero
        Console.WriteLine(Fix(42.9D) & "|" & Fix(-42.9D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42|-42"]);
}

#[test]
fn test_vb_decimal_int_function_behaviour() {
    let src = r#"
Module Program
    Sub Main()
        ' Int function rounds down towards negative infinity
        Console.WriteLine(Int(42.9D) & "|" & Int(-42.9D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42|-43"]);
}

#[test]
fn test_vb_decimal_sign_helper() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim pos As Decimal = 100D
        Dim neg As Decimal = -50D
        Dim zero As Decimal = 0D
        Console.WriteLine(Math.Sign(pos) & "|" & Math.Sign(neg) & "|" & Math.Sign(zero))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|-1|0"]);
}

#[test]
fn test_vb_decimal_negate_helper() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dec As Decimal = 42D
        Dim neg = Decimal.Negate(dec)
        Console.WriteLine(neg)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-42"]);
}

#[test]
fn test_vb_decimal_hash_code_equality() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 500.5D
        Dim d2 As Decimal = 500.5D
        Console.WriteLine(d1.GetHashCode() = d2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
