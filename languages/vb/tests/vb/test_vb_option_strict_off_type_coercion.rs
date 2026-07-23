use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Option Strict Off Type Coercions & Implicit Converts
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_option_strict_off_string_to_integer() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim strVal As Object = "100"
        Dim num As Integer = strVal
        Console.WriteLine(num + 50)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["150"]);
}

#[test]
fn test_vb_option_strict_off_integer_to_string_concat() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim num As Object = 42
        Dim str As String = num & " is the answer"
        Console.WriteLine(str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42 is the answer"]);
}

#[test]
fn test_vb_option_strict_off_double_to_integer_implicit_narrowing() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim dbl As Object = 12.8
        Dim intVal As Integer = dbl
        Console.WriteLine(intVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["13"]);
}

#[test]
fn test_vb_option_strict_off_boolean_to_integer_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim flag As Object = True
        Dim num As Integer = flag
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_option_strict_off_integer_to_boolean_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim zero As Object = 0
        Dim nonZero As Object = 5
        Dim b1 As Boolean = zero
        Dim b2 As Boolean = nonZero
        Console.WriteLine(b1 & "|" & b2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}

#[test]
fn test_vb_option_strict_off_string_addition_operator_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        ' With + operator, if left is numeric, string is coerced to numeric!
        Dim res = 10 + "20"
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_option_strict_off_string_subtraction_operator() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim res = "50" - "15"
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["35"]);
}

#[test]
fn test_vb_option_strict_off_string_multiplication_operator() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim res = "6" * "7"
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_option_strict_off_date_from_string() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim strDate As Object = "2025-01-01"
        Dim dt As DateTime = strDate
        Console.WriteLine(dt.Year)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025"]);
}

#[test]
fn test_vb_option_strict_off_decimal_from_string() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim strDec As Object = "99.99"
        Dim dec As Decimal = strDec
        Console.WriteLine(dec)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99.99"]);
}

#[test]
fn test_vb_option_strict_off_object_array_implicit_element_conversion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim objArr As Object() = {"10", "20", "30"}
        Dim first As Integer = objArr(0)
        Dim second As Integer = objArr(1)
        Console.WriteLine(first + second)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_option_strict_off_null_to_primitive_default() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Dim n As Integer = obj
        Dim b As Boolean = obj
        Dim s As String = obj
        Console.WriteLine(n & "|" & b & "|" & (s Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|False|True"]);
}

#[test]
fn test_vb_option_strict_off_enum_from_integer_or_string() {
    let src = r#"
Option Strict Off

Enum Status
    Inactive = 0
    Active = 1
End Enum

Module Program
    Sub Main()
        Dim obj1 As Object = 1
        Dim s1 As Status = obj1
        Console.WriteLine(s1.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Active"]);
}

#[test]
fn test_vb_option_strict_off_char_to_integer() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim objCh As Object = "A"c
        Dim charCode As Integer = objCh
        Console.WriteLine(charCode)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65"]);
}

#[test]
fn test_vb_option_strict_off_integer_to_char() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim code As Object = 66
        Dim ch As Char = code
        Console.WriteLine(ch)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B"]);
}

#[test]
fn test_vb_option_strict_off_invalid_string_conversion_throws() {
    let src = r#"
Imports System
Option Strict Off

Module Program
    Sub Main()
        Dim badStr As Object = "NotANumber"
        Try
            Dim n As Integer = badStr
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on String to Integer")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on String to Integer"]
    );
}

#[test]
fn test_vb_option_strict_off_comparison_operator_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim num As Object = 100
        Dim str As Object = "100"
        Console.WriteLine(num = str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_option_strict_off_bitwise_operator_string_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim res = "12" And "10" ' 1100 And 1010 = 1000 (8)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_option_strict_off_incompatible_reference_types_throws() {
    let src = r#"
Imports System
Option Strict Off

Class TypeA
End Class

Class TypeB
End Class

Module Program
    Sub Main()
        Dim a As Object = New TypeA()
        Try
            Dim b As TypeB = a
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on Incompatible Classes")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on Incompatible Classes"]
    );
}

#[test]
fn test_vb_option_strict_off_float_division_string_operands() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim res = "25" / "2"
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12.5"]);
}
