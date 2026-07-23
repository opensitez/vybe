use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Integer Overflow & Checked Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_overflow_integer_max_add() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim x As Integer = Integer.MaxValue
            Dim y As Integer = x + 1
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_integer_min_subtract() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim x As Integer = Integer.MinValue
            Dim y As Integer = x - 1
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_integer_multiply() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim x As Integer = 1000000
            Dim y As Integer = x * 1000000
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_short_max_add() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim x As Short = Short.MaxValue
            Dim y As Short = CShort(x + 1)
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_byte_max_add() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim b As Byte = Byte.MaxValue
            Dim b2 As Byte = CByte(b + 1)
            Console.WriteLine(b2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_long_max_add() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim l As Long = Long.MaxValue
            Dim l2 As Long = l + 1L
            Console.WriteLine(l2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_narrowing_double_to_integer() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim d As Double = 1e15
            Dim i As Integer = CInt(d)
            Console.WriteLine(i)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_negation_int_min() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim x As Integer = Integer.MinValue
            Dim y As Integer = -x
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_checked_math_abs_min() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim x As Integer = Integer.MinValue
            Dim y As Integer = Math.Abs(x)
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_unchecked_bitwise_shift_no_overflow() {
    let src = r#"
Module Program
    Sub Main()
        Dim x As Integer = 1
        Dim y As Integer = x Xor &H7FFFFFFF
        Console.WriteLine(y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2147483646"]);
}

#[test]
fn test_vb_overflow_ulong_max_add() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim ul As ULong = ULong.MaxValue
            Dim ul2 As ULong = ul + 1UL
            Console.WriteLine(ul2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_checked_convert_to_byte_negative() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim i As Integer = -1
            Dim b As Byte = Convert.ToByte(i)
            Console.WriteLine(b)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_checked_convert_to_sbyte_large() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim i As Integer = 200
            Dim sb As SByte = Convert.ToSByte(i)
            Console.WriteLine(sb)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_checked_decimal_max_multiply() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim d As Decimal = Decimal.MaxValue
            Dim d2 As Decimal = d * 2D
            Console.WriteLine(d2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_checked_checked_context_expr() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim a As Integer = Integer.MaxValue - 5
            Dim b As Integer = 10
            Dim c As Integer = a + b
            Console.WriteLine(c)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_safe_integer_addition() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Integer = 1000
        Dim b As Integer = 2000
        Dim c As Integer = a + b
        Console.WriteLine(c)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3000"]);
}

#[test]
fn test_vb_overflow_safe_long_addition() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Long = 5000000000L
        Dim b As Long = 5000000000L
        Dim c As Long = a + b
        Console.WriteLine(c)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10000000000"]);
}

#[test]
fn test_vb_overflow_checked_short_multiplication() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim s1 As Short = 1000
            Dim s2 As Short = 100
            Dim s3 As Short = CShort(s1 * s2)
            Console.WriteLine(s3)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}

#[test]
fn test_vb_overflow_checked_convert_to_uint16_negative() {
    let src = r#"
Module Program
    Sub Main()
        Try
            Dim i As Integer = -50
            Dim us As UShort = Convert.ToUInt16(i)
            Console.WriteLine(us)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException"]);
}
