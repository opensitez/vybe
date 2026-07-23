use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Numerics.BigInteger Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_bigint_parse_addition() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = BigInteger.Parse("123456789012345678901234567890")
        Dim b2 As BigInteger = BigInteger.Parse("987654321098765432109876543210")
        Dim sum As BigInteger = b1 + b2
        Console.WriteLine(sum.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1111111110111111111011111111100"]);
}

#[test]
fn test_vb_bigint_multiplication_factorial() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim fact As BigInteger = 1
        For i As Integer = 1 To 20
            fact *= i
        Next
        Console.WriteLine(fact.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2432902008176640000"]);
}

#[test]
fn test_vb_bigint_subtraction_negative() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = 100
        Dim b2 As BigInteger = 500
        Dim diff As BigInteger = b1 - b2
        Console.WriteLine(diff.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-400"]);
}

#[test]
fn test_vb_bigint_division_and_modulus() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = 1000
        Dim b2 As BigInteger = 300
        Dim div As BigInteger = b1 \ b2
        Dim remVal As BigInteger = b1 Mod b2
        Console.WriteLine(div.ToString())
        Console.WriteLine(remVal.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "100"]);
}

#[test]
fn test_vb_bigint_pow_function() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b As BigInteger = 2
        Dim pow As BigInteger = BigInteger.Pow(b, 64)
        Console.WriteLine(pow.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["18446744073709551616"]);
}

#[test]
fn test_vb_bigint_greatest_common_divisor() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = 54
        Dim b2 As BigInteger = 24
        Dim gcd As BigInteger = BigInteger.GreatestCommonDivisor(b1, b2)
        Console.WriteLine(gcd.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_bigint_bitwise_and_or_xor() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = &HFF00
        Dim b2 As BigInteger = &H00FF
        Dim andRes As BigInteger = b1 And b2
        Dim orRes As BigInteger = b1 Or b2
        Console.WriteLine(andRes.ToString())
        Console.WriteLine(orRes.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "65535"]);
}

#[test]
fn test_vb_bigint_zero_one_minus_one_properties() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Console.WriteLine(BigInteger.Zero.ToString())
        Console.WriteLine(BigInteger.One.ToString())
        Console.WriteLine(BigInteger.MinusOne.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "1", "-1"]);
}

#[test]
fn test_vb_bigint_is_power_of_two() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = 1024
        Dim b2 As BigInteger = 1000
        Console.WriteLine(b1.IsPowerOfTwo)
        Console.WriteLine(b2.IsPowerOfTwo)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_bigint_is_even_is_one() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = 42
        Dim b2 As BigInteger = 1
        Console.WriteLine(b1.IsEven)
        Console.WriteLine(b2.IsOne)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_bigint_to_byte_array() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b As BigInteger = 258
        Dim bytes As Byte() = b.ToByteArray()
        Console.WriteLine(bytes.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_bigint_from_byte_array() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim bytes As Byte() = {2, 1, 0}
        Dim b As BigInteger = New BigInteger(bytes)
        Console.WriteLine(b.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["258"]);
}

#[test]
fn test_vb_bigint_compare_to() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = 100
        Dim b2 As BigInteger = 200
        Console.WriteLine(b1.CompareTo(b2) < 0)
        Console.WriteLine(b2.CompareTo(b1) > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_bigint_explicit_conversion_from_long() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim l As Long = 9223372036854775807L
        Dim b As BigInteger = l
        Console.WriteLine(b.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9223372036854775807"]);
}

#[test]
fn test_vb_bigint_explicit_conversion_to_int() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b As BigInteger = 123456
        Dim i As Integer = CInt(b)
        Console.WriteLine(i)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123456"]);
}

#[test]
fn test_vb_bigint_abs_negate() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b As BigInteger = -999
        Dim absVal As BigInteger = BigInteger.Abs(b)
        Dim negVal As BigInteger = BigInteger.Negate(absVal)
        Console.WriteLine(absVal.ToString())
        Console.WriteLine(negVal.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999", "-999"]);
}

#[test]
fn test_vb_bigint_log_log10() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b As BigInteger = 1000
        Dim log10Val As Double = BigInteger.Log10(b)
        Console.WriteLine(log10Val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_bigint_tryparse_valid() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim res As BigInteger
        Dim ok As Boolean = BigInteger.TryParse("99999999999999999999", res)
        Console.WriteLine(ok)
        Console.WriteLine(res.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "99999999999999999999"]);
}
