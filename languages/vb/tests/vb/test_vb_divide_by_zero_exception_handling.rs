use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DivideByZeroException & Floating Point Arithmetic
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_integer_division_by_zero_throws_divide_by_zero_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim a As Integer = 10
            Dim b As Integer = 0
            Dim res As Integer = a \ b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("DivideByZeroException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DivideByZeroException Caught"]);
}

#[test]
fn test_vb_integer_mod_by_zero_throws_divide_by_zero_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim a As Integer = 10
            Dim b As Integer = 0
            Dim res As Integer = a Mod b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Mod DivideByZeroException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Mod DivideByZeroException Caught"]);
}

#[test]
fn test_vb_floating_point_division_by_zero_returns_infinity() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Double = 10.0
        Dim b As Double = 0.0
        Dim res As Double = a / b
        Console.WriteLine(Double.IsInfinity(res) & "|" & (res > 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_floating_point_negative_division_by_zero_returns_negative_infinity() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Double = -10.0
        Dim b As Double = 0.0
        Dim res As Double = a / b
        Console.WriteLine(Double.IsNegativeInfinity(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_zero_divided_by_zero_floating_point_returns_nan() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Double = 0.0
        Dim b As Double = 0.0
        Dim res As Double = a / b
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_decimal_division_by_zero_throws_divide_by_zero_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim a As Decimal = 10.5D
            Dim b As Decimal = 0D
            Dim res As Decimal = a / b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Decimal DivideByZeroException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Decimal DivideByZeroException Caught"]);
}

#[test]
fn test_vb_long_integer_division_by_zero() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim a As Long = 1000000000000L
            Dim b As Long = 0L
            Dim res As Long = a \ b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Long DivideByZeroException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Long DivideByZeroException Caught"]);
}

#[test]
fn test_vb_short_integer_division_by_zero() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim a As Short = 100S
            Dim b As Short = 0S
            Dim res As Short = CShort(a \ b)
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Short DivideByZeroException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Short DivideByZeroException Caught"]);
}

#[test]
fn test_vb_byte_integer_division_by_zero() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim a As Byte = 255
            Dim b As Byte = 0
            Dim res As Byte = CByte(a \ b)
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Byte DivideByZeroException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Byte DivideByZeroException Caught"]);
}

#[test]
fn test_vb_single_float_division_by_zero_returns_positive_infinity() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Single = 5.0F
        Dim b As Single = 0.0F
        Dim res As Single = a / b
        Console.WriteLine(Single.IsPositiveInfinity(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_single_float_zero_by_zero_returns_nan() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Single = 0.0F
        Dim b As Single = 0.0F
        Dim res As Single = a / b
        Console.WriteLine(Single.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_slash_operator_integers_promotes_to_double() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Integer = 10
        Dim b As Integer = 0
        ' Visual Basic "/" operator between integers performs floating point division!
        Dim res As Double = a / b
        Console.WriteLine(Double.IsInfinity(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_backslash_operator_integer_division_strict() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim a As Integer = 10
            Dim b As Integer = 0
            ' Visual Basic "\" operator performs strict integer division and throws!
            Dim res As Integer = a \ b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Backslash Throws DivideByZeroException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Backslash Throws DivideByZeroException"]);
}

#[test]
fn test_vb_divide_by_zero_checked_block() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim x As Integer = 50
            Dim y As Integer = 0
            Dim z As Integer = x \ y
        Catch ex As DivideByZeroException
            Console.WriteLine("Checked DivideByZeroException Handled")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Checked DivideByZeroException Handled"]);
}

#[test]
fn test_vb_divide_by_zero_in_compound_assignment() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim x As Integer = 100
            Dim zero As Integer = 0
            x \= zero
        Catch ex As DivideByZeroException
            Console.WriteLine("Compound Backslash DivideByZeroException Handled")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Compound Backslash DivideByZeroException Handled"]
    );
}

#[test]
fn test_vb_divide_by_zero_in_expression_chain() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim val As Integer = (100 + 50) \ (10 - 10)
        Catch ex As DivideByZeroException
            Console.WriteLine("Chain DivideByZeroException Handled")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Chain DivideByZeroException Handled"]);
}

#[test]
fn test_vb_divide_by_zero_in_linq_query() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 0, 30}
        Try
            Dim results = (From n In numbers Select 100 \ n).ToList()
        Catch ex As DivideByZeroException
            Console.WriteLine("LINQ DivideByZeroException Handled")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LINQ DivideByZeroException Handled"]);
}

#[test]
fn test_vb_infinity_multiplication_with_zero_is_nan() {
    let src = r#"
Module Program
    Sub Main()
        Dim inf As Double = Double.PositiveInfinity
        Dim zero As Double = 0.0
        Dim result As Double = inf * zero
        Console.WriteLine(Double.IsNaN(result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_nan_comparison_behavior() {
    let src = r#"
Module Program
    Sub Main()
        Dim nan As Double = Double.NaN
        Console.WriteLine((nan = nan) & "|" & Double.IsNaN(nan))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}

#[test]
fn test_vb_divide_by_zero_custom_helper_fallback() {
    let src = r#"
Imports System

Module Program
    Private Function SafeDivide(a As Integer, b As Integer) As Integer
        Try
            Return a \ b
        Catch ex As DivideByZeroException
            Return 0
        End Try
    End Function

    Sub Main()
        Console.WriteLine(SafeDivide(10, 2) & "|" & SafeDivide(10, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|0"]);
}
