use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Decimal Arithmetic Precision
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_decimal_exact_addition() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 0.1D
        Dim d2 As Decimal = 0.2D
        Dim sum As Decimal = d1 + d2
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0.3"]);
}

#[test]
fn test_vb_decimal_exact_subtraction() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 1.0D
        Dim d2 As Decimal = 0.99999999999999999999D
        Dim diff As Decimal = d1 - d2
        Console.WriteLine(diff)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0.00000000000000000001"]);
}

#[test]
fn test_vb_decimal_multiplication_precision() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 123.456D
        Dim d2 As Decimal = 789.012D
        Dim prod As Decimal = d1 * d2
        Console.WriteLine(prod)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["97408.261472"]);
}

#[test]
fn test_vb_decimal_division_truncation() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 1D
        Dim d2 As Decimal = 3D
        Dim div As Decimal = d1 / d2
        Console.WriteLine(div.ToString().Length > 20)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_decimal_remainder_operator() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 10.5D
        Dim d2 As Decimal = 3.0D
        Dim remVal As Decimal = d1 Mod d2
        Console.WriteLine(remVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.5"]);
}

#[test]
fn test_vb_decimal_round_banker_even() {
    let src = r#"
Module Program
    Sub Main()
        Dim r1 As Decimal = Decimal.Round(2.5D, 0, MidpointRounding.ToEven)
        Dim r2 As Decimal = Decimal.Round(3.5D, 0, MidpointRounding.ToEven)
        Console.WriteLine(r1)
        Console.WriteLine(r2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "4"]);
}

#[test]
fn test_vb_decimal_round_away_from_zero() {
    let src = r#"
Module Program
    Sub Main()
        Dim r1 As Decimal = Decimal.Round(2.5D, 0, MidpointRounding.AwayFromZero)
        Dim r2 As Decimal = Decimal.Round(3.5D, 0, MidpointRounding.AwayFromZero)
        Console.WriteLine(r1)
        Console.WriteLine(r2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "4"]);
}

#[test]
fn test_vb_decimal_truncate_function() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = Decimal.Truncate(12.89D)
        Dim d2 As Decimal = Decimal.Truncate(-12.89D)
        Console.WriteLine(d1)
        Console.WriteLine(d2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12", "-12"]);
}

#[test]
fn test_vb_decimal_floor_function() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = Decimal.Floor(12.89D)
        Dim d2 As Decimal = Decimal.Floor(-12.89D)
        Console.WriteLine(d1)
        Console.WriteLine(d2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12", "-13"]);
}

#[test]
fn test_vb_decimal_ceiling_function() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = Decimal.Ceiling(12.01D)
        Dim d2 As Decimal = Decimal.Ceiling(-12.89D)
        Console.WriteLine(d1)
        Console.WriteLine(d2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["13", "-12"]);
}

#[test]
fn test_vb_decimal_get_bits_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim d As Decimal = 10.0D
        Dim bits As Integer() = Decimal.GetBits(d)
        Console.WriteLine(bits.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_decimal_from_bits_reconstruct() {
    let src = r#"
Module Program
    Sub Main()
        Dim bits As Integer() = {100, 0, 0, 0}
        Dim d As Decimal = New Decimal(bits)
        Console.WriteLine(d)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_decimal_negation_operator() {
    let src = r#"
Module Program
    Sub Main()
        Dim d As Decimal = 45.67D
        Dim neg As Decimal = -d
        Console.WriteLine(neg)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-45.67"]);
}

#[test]
fn test_vb_decimal_min_max_constants() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(Decimal.MinValue < 0D)
        Console.WriteLine(Decimal.MaxValue > 0D)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_decimal_one_zero_minus_one() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(Decimal.Zero)
        Console.WriteLine(Decimal.One)
        Console.WriteLine(Decimal.MinusOne)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "1", "-1"]);
}

#[test]
fn test_vb_decimal_parse_valid() {
    let src = r#"
Module Program
    Sub Main()
        Dim d As Decimal = Decimal.Parse("9876543210.123456789")
        Console.WriteLine(d)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9876543210.123456789"]);
}

#[test]
fn test_vb_decimal_tryparse_valid() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As Decimal
        Dim ok As Boolean = Decimal.TryParse("123.456", res)
        Console.WriteLine(ok)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "123.456"]);
}

#[test]
fn test_vb_decimal_comparison_equality() {
    let src = r#"
Module Program
    Sub Main()
        Dim d1 As Decimal = 10.00D
        Dim d2 As Decimal = 10.0D
        Console.WriteLine(d1 = d2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_decimal_conversion_to_double() {
    let src = r#"
Module Program
    Sub Main()
        Dim d As Decimal = 123.45D
        Dim dbl As Double = CDbl(d)
        Console.WriteLine(dbl)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123.45"]);
}
