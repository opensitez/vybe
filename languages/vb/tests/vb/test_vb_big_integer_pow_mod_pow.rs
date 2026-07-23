use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Numerics.BigInteger Powers & ModPow Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_big_integer_pow_large_exponent() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim baseVal As New BigInteger(2)
        Dim res = BigInteger.Pow(baseVal, 32)
        Console.WriteLine(res.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4294967296"]);
}

#[test]
fn test_vb_big_integer_mod_pow_crypto_calculation() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim value As New BigInteger(10)
        Dim exponent As New BigInteger(3)
        Dim modulus As New BigInteger(7)
        ' (10^3) mod 7 = 1000 mod 7 = 6
        Dim res = BigInteger.ModPow(value, exponent, modulus)
        Console.WriteLine(res.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_big_integer_greatest_common_divisor() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim a As New BigInteger(54)
        Dim b As New BigInteger(24)
        Dim gcd = BigInteger.GreatestCommonDivisor(a, b)
        Console.WriteLine(gcd.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_big_integer_log10_log_natural() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As New BigInteger(1000)
        Dim log10Val = BigInteger.Log10(val)
        Console.WriteLine(log10Val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_big_integer_parse_hex_string() {
    let src = r#"
Imports System.Globalization
Imports System.Numerics

Module Program
    Sub Main()
        Dim hexStr = "00FF"
        Dim val = BigInteger.Parse(hexStr, NumberStyles.HexNumber)
        Console.WriteLine(val.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_big_integer_bitwise_and_or_xor() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim a As New BigInteger(12) ' 1100
        Dim b As New BigInteger(10) ' 1010
        Dim andVal = a And b ' 1000 (8)
        Dim orVal = a Or b   ' 1110 (14)
        Dim xorVal = a Xor b ' 0110 (6)
        Console.WriteLine(andVal.ToString() & "|" & orVal.ToString() & "|" & xorVal.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8|14|6"]);
}

#[test]
fn test_vb_big_integer_left_shift_right_shift() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As New BigInteger(1)
        Dim shiftedLeft = val << 10 ' 1024
        Dim shiftedRight = shiftedLeft >> 5 ' 32
        Console.WriteLine(shiftedLeft.ToString() & "|" & shiftedRight.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1024|32"]);
}

#[test]
fn test_vb_big_integer_arithmetic_operators() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim a As BigInteger = 1000000000000000000L
        Dim b As BigInteger = 2000000000000000000L
        Dim sum = a + b
        Dim diff = b - a
        Dim mult = a * 2
        Console.WriteLine(sum.ToString() & "|" & diff.ToString() & "|" & mult.ToString())
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["3000000000000000000|1000000000000000000|2000000000000000000"]
    );
}

#[test]
fn test_vb_big_integer_negation_operator() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim positive As New BigInteger(500)
        Dim negative = -positive
        Dim absVal = BigInteger.Abs(negative)
        Console.WriteLine(negative.ToString() & "|" & absVal.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-500|500"]);
}

#[test]
fn test_vb_big_integer_comparison_operators() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim a As BigInteger = 100
        Dim b As BigInteger = 200
        Console.WriteLine((a < b) & "|" & (a <= b) & "|" & (b > a) & "|" & (a = b) & "|" & (a <> b))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True|False|True"]);
}

#[test]
fn test_vb_big_integer_is_even_is_power_of_two() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As New BigInteger(1024)
        Console.WriteLine(val.IsEven & "|" & val.IsPowerOfTwo)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_big_integer_is_zero_is_one() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim zero = BigInteger.Zero
        Dim one = BigInteger.One
        Dim minusOne = BigInteger.MinusOne
        Console.WriteLine(zero.IsZero & "|" & one.IsOne & "|" & (minusOne.Sign = -1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_big_integer_to_byte_array_roundtrip() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim original As New BigInteger(1234567890)
        Dim bytes = original.ToByteArray()
        Dim restored As New BigInteger(bytes)
        Console.WriteLine(original = restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_big_integer_implicit_explicit_casts() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim big: BigInteger = 999
        Dim intVal As Integer = CInt(big)
        Dim doubleVal As Double = CDbl(big)
        Console.WriteLine(intVal & "|" & doubleVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999|999"]);
}

#[test]
fn test_vb_big_integer_min_max_static_helpers() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim a As BigInteger = 50
        Dim b As BigInteger = 100
        Console.WriteLine(BigInteger.Min(a, b).ToString() & "|" & BigInteger.Max(a, b).ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50|100"]);
}

#[test]
fn test_vb_big_integer_div_rem_tuple() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim dividend As New BigInteger(25)
        Dim divisor As New BigInteger(7)
        Dim remainder As BigInteger
        Dim quotient = BigInteger.DivRem(dividend, divisor, remainder)
        Console.WriteLine(quotient.ToString() & " R " & remainder.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3 R 4"]);
}

#[test]
fn test_vb_big_integer_clamp_helper() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As BigInteger = 150
        Dim min As BigInteger = 0
        Dim max As BigInteger = 100
        Dim clamped = BigInteger.Clamp(val, min, max)
        Console.WriteLine(clamped.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_big_integer_tostring_formatting_hex() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As New BigInteger(255)
        Console.WriteLine(val.ToString("X"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FF"]);
}

#[test]
fn test_vb_big_integer_factorial_simulation() {
    let src = r#"
Imports System.Numerics

Module Program
    Private Function Factorial(n As Integer) As BigInteger
        Dim result As BigInteger = 1
        For i As Integer = 2 To n
            result *= i
        Next
        Return result
    End Function

    Sub Main()
        Dim fact20 = Factorial(20)
        Console.WriteLine(fact20.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2432902008176640000"]);
}

#[test]
fn test_vb_big_integer_equality_hashcode() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim a As BigInteger = 1000
        Dim b As BigInteger = 1000
        Console.WriteLine(a.Equals(b) & "|" & (a.GetHashCode() = b.GetHashCode()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}
