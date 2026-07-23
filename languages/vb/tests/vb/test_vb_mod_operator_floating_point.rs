use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Mod Operator Mechanics across Integer & Floating Point
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_mod_operator_positive_integers() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = 17 Mod 5
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_mod_operator_negative_dividend() {
    let src = r#"
Module Program
    Sub Main()
        ' In VB.NET Mod: sign of result matches sign of dividend! -17 Mod 5 = -2
        Dim res = -17 Mod 5
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-2"]);
}

#[test]
fn test_vb_mod_operator_negative_divisor() {
    let src = r#"
Module Program
    Sub Main()
        ' In VB.NET Mod: 17 Mod -5 = 2
        Dim res = 17 Mod -5
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_mod_operator_both_negative() {
    let src = r#"
Module Program
    Sub Main()
        ' -17 Mod -5 = -2
        Dim res = -17 Mod -5
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-2"]);
}

#[test]
fn test_vb_mod_operator_floating_point_double() {
    let src = r#"
Module Program
    Sub Main()
        ' Unlike C#, VB.NET Mod rounds floating point operands to Long before computing Mod!
        Dim a As Double = 17.6 ' Rounds to 18
        Dim b As Double = 4.9  ' Rounds to 5
        Dim res = a Mod b      ' 18 Mod 5 = 3
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_mod_operator_decimal_operands() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Decimal = 10.5D
        Dim b As Decimal = 3D
        Dim res = a Mod b
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.5"]);
}

#[test]
fn test_vb_mod_operator_byte_operands() {
    let src = r#"
Module Program
    Sub Main()
        Dim b1 As Byte = 250
        Dim b2 As Byte = 40
        Dim res As Byte = CByte(b1 Mod b2)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_mod_operator_long_operands() {
    let src = r#"
Module Program
    Sub Main()
        Dim l1 As Long = 10000000000L
        Dim l2 As Long = 3000000000L
        Dim res = l1 Mod l2
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1000000000"]);
}

#[test]
fn test_vb_mod_operator_divide_by_zero_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim res = 10 Mod 0
        Catch ex As DivideByZeroException
            Console.WriteLine("DivideByZeroException Caught on Mod 0")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DivideByZeroException Caught on Mod 0"]);
}

#[test]
fn test_vb_mod_operator_zero_dividend() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = 0 Mod 5
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_mod_operator_exact_multiple() {
    let src = r#"
Module Program
    Sub Main()
        Dim res = 100 Mod 25
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_mod_operator_even_odd_checker() {
    let src = r#"
Module Program
    Private Function IsEven(n As Integer) As Boolean
        Return (n Mod 2) = 0
    End Function

    Sub Main()
        Console.WriteLine(IsEven(10) & "|" & IsEven(11))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_mod_operator_compound_assignment() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 27
        val %= 5
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_mod_operator_uinteger_operands() {
    let src = r#"
Module Program
    Sub Main()
        Dim u1 As UInteger = 4000000000UI
        Dim u2 As UInteger = 3UI
        Dim res = u1 Mod u2
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_mod_operator_ulong_operands() {
    let src = r#"
Module Program
    Sub Main()
        Dim u1 As ULong = 18000000000000000000UL
        Dim u2 As ULong = 7UL
        Dim res = u1 Mod u2
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_mod_operator_string_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim res = "29" Mod "6"
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_mod_operator_bankers_rounding_float_operands() {
    let src = r#"
Module Program
    Sub Main()
        ' 2.5 rounds to 2 (even), 3.5 rounds to 4 (even)
        ' 2 Mod 4 = 2
        Dim res = 2.5 Mod 3.5
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_mod_operator_operator_precedence() {
    let src = r#"
Module Program
    Sub Main()
        ' Addition has lower precedence than Mod! 10 + 17 Mod 5 = 10 + 2 = 12
        Dim res = 10 + 17 Mod 5
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12"]);
}

#[test]
fn test_vb_mod_operator_biginteger_type() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = 100000000000000000000000000000000000000D
        Dim b2 As BigInteger = 3
        Dim res = b1 Mod b2
        Console.WriteLine(res.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_mod_operator_custom_class_operator_overload() {
    let src = r#"
Module Program
    Class ClockTime
        Public Hours As Integer
        Public Sub New(h As Integer)
            Hours = h
        End Sub
        Public Shared Operator Mod(a As ClockTime, b As Integer) As ClockTime
            Return New ClockTime(a.Hours Mod b)
        End Operator
    End Class

    Sub Main()
        Dim t As New ClockTime(27)
        Dim wrapped = t Mod 24
        Console.WriteLine(wrapped.Hours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}
