use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Option Strict On/Off Implicit Coercions & Typing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_option_strict_off_implicit_string_to_integer() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim strVal As String = "100"
        Dim num As Integer = strVal ' Implicit coercion under Option Strict Off
        Console.WriteLine(num + 50)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["150"]);
}

#[test]
fn test_vb_option_strict_off_implicit_double_to_integer_truncation() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim dblVal As Double = 99.8
        Dim num As Integer = dblVal ' Implicit narrowing under Option Strict Off (rounds to nearest even)
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_option_strict_off_late_bound_method_call() {
    let src = r#"
Option Strict Off

Class DynamicWorker
    Public Function Work(msg As String) As String
        Return "Worked: " & msg
    End Function
End Class

Module Program
    Sub Main()
        Dim obj As Object = New DynamicWorker()
        ' Late bound call without explicit cast under Option Strict Off
        Dim res = obj.Work("Task1")
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Worked: Task1"]);
}

#[test]
fn test_vb_option_strict_off_implicit_boolean_to_string() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim flag As Boolean = True
        Dim msg As String = flag
        Console.WriteLine("[" & msg & "]")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[True]"]);
}

#[test]
fn test_vb_option_strict_on_explicit_cast_required() {
    let src = r#"
Option Strict On

Module Program
    Sub Main()
        Dim dbl As Double = 12.34
        ' Requires explicit CInt under Option Strict On
        Dim num As Integer = CInt(dbl)
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12"]);
}

#[test]
fn test_vb_option_strict_on_widening_conversions_implicit_allowed() {
    let src = r#"
Option Strict On

Module Program
    Sub Main()
        Dim b As Byte = 255
        Dim i As Integer = b ' Widening Byte to Integer allowed under Option Strict On
        Dim d As Double = i  ' Widening Integer to Double allowed
        Console.WriteLine(d)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_option_strict_off_implicit_array_element_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim rawObjs As Object() = New Object() {"10", "20", "30"}
        Dim firstNum As Integer = rawObjs(0)
        Console.WriteLine(firstNum * 2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_option_strict_off_binary_string_addition_concatenation() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        ' String + String concatenates, but String + Integer coerces String to Integer!
        Dim res1 = "10" + "20"
        Dim res2 = "10" + 20
        Console.WriteLine(res1 & "|" & res2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1020|30"]);
}

#[test]
fn test_vb_option_strict_off_late_bound_property_assignment() {
    let src = r#"
Option Strict Off

Class Config
    Public Property Value As String
End Class

Module Program
    Sub Main()
        Dim cfg As Object = New Config()
        cfg.Value = 12345 ' Numeric assigned to String property late-bound
        Console.WriteLine(cfg.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12345"]);
}

#[test]
fn test_vb_option_strict_off_implicit_date_time_parsing() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim dateStr As String = "2025-01-01"
        Dim dt As DateTime = dateStr
        Console.WriteLine(dt.Year)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025"]);
}

#[test]
fn test_vb_option_strict_off_implicit_enum_conversion() {
    let src = r#"
Option Strict Off

Enum Mode
    Off = 0
    OnVal = 1
End Enum

Module Program
    Sub Main()
        Dim num As Integer = 1
        Dim m As Mode = num ' Implicit conversion from Integer to Enum
        Console.WriteLine(m.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OnVal"]);
}

#[test]
fn test_vb_option_strict_off_char_to_string_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim ch As Char = "Z"c
        Dim s As String = ch
        Console.WriteLine(s & ":" & s.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Z:1"]);
}

#[test]
fn test_vb_option_strict_off_object_relational_comparison() {
    let src = r#"
Option Strict Off

Module Program
    Sub Main()
        Dim o1 As Object = "50"
        Dim o2 As Object = 20
        ' Coerces "50" to numeric 50 and compares > 20
        Dim isGreater = o1 > o2
        Console.WriteLine(isGreater)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_option_strict_on_lambdas_require_typed_parameters() {
    let src = r#"
Option Strict On
Imports System

Module Program
    Sub Main()
        ' In Option Strict On lambda params require explicit types or inferred from delegate!
        Dim sq As Func(Of Integer, Integer) = Function(x As Integer) x * x
        Console.WriteLine(sq(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["25"]);
}

#[test]
fn test_vb_option_strict_off_invalid_coercion_runtime_error() {
    let src = r#"
Option Strict Off
Imports System

Module Program
    Sub Main()
        Dim invalidStr As String = "ABC"
        Try
            Dim num As Integer = invalidStr ' Runtime failure when coercion fails!
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on Coercion")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["InvalidCastException Caught on Coercion"]);
}

#[test]
fn test_vb_option_strict_off_overload_resolution_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Private Sub PrintValue(n As Integer)
        Console.WriteLine("Int: " & n)
    End Sub

    Sub Main()
        PrintValue("99") ' String coerced to Int parameter
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int: 99"]);
}

#[test]
fn test_vb_option_strict_off_indexed_object_late_binding() {
    let src = r#"
Option Strict Off
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As Object = New Dictionary(Of String, Integer)()
        dict.Add("Key1", 500)
        Console.WriteLine(dict("Key1"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["500"]);
}

#[test]
fn test_vb_option_strict_on_derived_class_to_base_implicit() {
    let src = r#"
Option Strict On

Class Base
End Class

Class Derived
    Inherits Base
End Class

Module Program
    Sub Main()
        Dim d As New Derived()
        Dim b As Base = d ' Reference widening is allowed under Option Strict On!
        Console.WriteLine(b IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_option_strict_off_byref_parameter_coercion() {
    let src = r#"
Option Strict Off

Module Program
    Private Sub DoubleVal(ByRef x As Integer)
        x *= 2
    End Sub

    Sub Main()
        Dim s As String = "15"
        DoubleVal(s) ' Coerced ByRef parameter
        Console.WriteLine(s)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_option_strict_off_math_pow_coercion() {
    let src = r#"
Option Strict Off
Imports System

Module Program
    Sub Main()
        Dim baseVal As String = "2"
        Dim expVal As String = "10"
        Dim res = Math.Pow(baseVal, expVal)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1024"]);
}
