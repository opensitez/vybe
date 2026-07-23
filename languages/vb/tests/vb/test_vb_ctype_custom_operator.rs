use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: CType Operator & Custom Widening/Narrowing Conversion Operators
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_ctype_custom_widening_operator() {
    let src = r#"
Imports System

Class Distance
    Public Meters As Double
    Public Sub New(m As Double)
        Meters = m
    End Sub

    ' Widening operator: Double to Distance
    Public Shared Widening Operator CType(m As Double) As Distance
        Return New Distance(m)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        ' Explicit or implicit CType call uses Widening operator!
        Dim d As Distance = CType(100.5, Distance)
        Console.WriteLine(d.Meters)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100.5"]);
}

#[test]
fn test_vb_ctype_custom_narrowing_operator() {
    let src = r#"
Imports System

Class Temperature
    Public Celsius As Double
    Public Sub New(c As Double)
        Celsius = c
    End Sub

    ' Narrowing operator: Temperature to Integer (may lose decimal precision)
    Public Shared Narrowing Operator CType(t As Temperature) As Integer
        Return CInt(t.Celsius)
    End Shared Narrowing Operator
End Class

Module Program
    Sub Main()
        Dim temp As New Temperature(36.6)
        Dim cInt As Integer = CType(temp, Integer)
        Console.WriteLine(cInt)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["37"]);
}

#[test]
fn test_vb_ctype_bidirectional_conversion_operators() {
    let src = r#"
Class Celsius
    Public Degrees As Double
    Public Sub New(d As Double)
        Degrees = d
    End Sub
End Class

Class Fahrenheit
    Public Degrees As Double
    Public Sub New(d As Double)
        Degrees = d
    End Sub

    Public Shared Widening Operator CType(c As Celsius) As Fahrenheit
        Return New Fahrenheit(c.Degrees * 9.0 / 5.0 + 32.0)
    End Shared Widening Operator

    Public Shared Widening Operator CType(f As Fahrenheit) As Celsius
        Return New Celsius((f.Degrees - 32.0) * 5.0 / 9.0)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim c As New Celsius(100)
        Dim f As Fahrenheit = CType(c, Fahrenheit)
        Dim restoredC As Celsius = CType(f, Celsius)
        Console.WriteLine(f.Degrees & "|" & restoredC.Degrees)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["212|100"]);
}

#[test]
fn test_vb_ctype_structure_to_primitive_conversion() {
    let src = r#"
Structure ComplexNumber
    Public Real As Double
    Public Imaginary As Double
    Public Sub New(r As Double, i As Double)
        Real = r
        Imaginary = i
    End Sub

    Public Shared Narrowing Operator CType(c As ComplexNumber) As Double
        Return c.Real
    End Shared Narrowing Operator
End Structure

Module Program
    Sub Main()
        Dim c As New ComplexNumber(42.5, 3.0)
        Dim r As Double = CType(c, Double)
        Console.WriteLine(r)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42.5"]);
}

#[test]
fn test_vb_ctype_primitive_widening_conversions() {
    let src = r#"
Module Program
    Sub Main()
        Dim b As Byte = 200
        Dim i As Integer = CType(b, Integer)
        Dim d As Double = CType(i, Double)
        Dim dec As Decimal = CType(d, Decimal)
        Console.WriteLine(dec)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200"]);
}

#[test]
fn test_vb_ctype_string_to_primitive_conversions() {
    let src = r#"
Module Program
    Sub Main()
        Dim numStr = "12345"
        Dim dblStr = "99.9"
        Dim num As Integer = CType(numStr, Integer)
        Dim dbl As Double = CType(dblStr, Double)
        Console.WriteLine(num + 5 & "|" & dbl)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12350|99.9"]);
}

#[test]
fn test_vb_ctype_enum_to_integer_conversion() {
    let src = r#"
Enum Level
    Low = 10
    High = 90
End Enum

Module Program
    Sub Main()
        Dim l = Level.High
        Dim val As Integer = CType(l, Integer)
        Dim restored As Level = CType(val, Level)
        Console.WriteLine(val & "|" & restored.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["90|High"]);
}

#[test]
fn test_vb_ctype_custom_operator_throws_overflow_exception() {
    let src = r#"
Imports System

Class BoundedVal
    Public Value As Long
    Public Sub New(v As Long)
        Value = v
    End Sub

    Public Shared Narrowing Operator CType(b As BoundedVal) As Byte
        If b.Value < 0 OrElse b.Value > 255 Then Throw New OverflowException("BoundedVal Byte Overflow")
        Return CByte(b.Value)
    End Shared Narrowing Operator
End Class

Module Program
    Sub Main()
        Dim bv As New BoundedVal(1000)
        Try
            Dim b As Byte = CType(bv, Byte)
        Catch ex As OverflowException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BoundedVal Byte Overflow"]);
}

#[test]
fn test_vb_ctype_interface_implementation_casting() {
    let src = r#"
Imports System

Interface IStorable
    Sub Save()
End Interface

Class Document
    Implements IStorable
    Public Sub Save() Implements IStorable.Save
        Console.WriteLine("Document Saved")
    End Sub
End Class

Module Program
    Sub Main()
        Dim doc As New Document()
        Dim storable As IStorable = CType(doc, IStorable)
        storable.Save()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Document Saved"]);
}

#[test]
fn test_vb_ctype_generic_struct_conversion_operator() {
    let src = r#"
Structure Wrapper(Of T)
    Public Value As T
    Public Sub New(v As T)
        Value = v
    End Sub

    Public Shared Widening Operator CType(v As T) As Wrapper(Of T)
        Return New Wrapper(Of T)(v)
    End Shared Widening Operator
End Structure

Module Program
    Sub Main()
        Dim w As Wrapper(Of String) = CType("WrappedString", Wrapper(Of String))
        Console.WriteLine(w.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["WrappedString"]);
}

#[test]
fn test_vb_ctype_array_to_list_conversion() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim arr As Integer() = {1, 2, 3}
        Dim list As New List(Of Integer)(arr)
        Console.WriteLine(String.Join("-", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1-2-3"]);
}

#[test]
fn test_vb_ctype_char_to_integer_ascii_code() {
    let src = r#"
Module Program
    Sub Main()
        Dim ch As Char = "A"c
        Dim code As Integer = CType(ch, Integer)
        Dim restored As Char = CType(code, Char)
        Console.WriteLine(code & "|" & restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65|A"]);
}

#[test]
fn test_vb_ctype_boolean_to_integer_conversion() {
    let src = r#"
Module Program
    Sub Main()
        ' In VB.NET CType(True, Integer) = -1, CType(False, Integer) = 0!
        Dim tVal As Integer = CType(True, Integer)
        Dim fVal As Integer = CType(False, Integer)
        Console.WriteLine(tVal & "|" & fVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1|0"]);
}

#[test]
fn test_vb_ctype_integer_to_boolean_conversion() {
    let src = r#"
Module Program
    Sub Main()
        Dim b1 As Boolean = CType(-1, Boolean)
        Dim b2 As Boolean = CType(100, Boolean)
        Dim b3 As Boolean = CType(0, Boolean)
        Console.WriteLine(b1 & "|" & b2 & "|" & b3)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_ctype_null_object_to_value_type_returns_default() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Dim n As Integer = CType(obj, Integer)
        Console.WriteLine(n)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_ctype_date_time_to_string_conversion() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 12, 25)
        Dim s As String = CType(dt, String)
        Console.WriteLine(s.Contains("2025"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_ctype_custom_operator_inheritance_lookup() {
    let src = r#"
Class BaseVal
    Public X As Integer
    Public Sub New(val As Integer)
        X = val
    End Sub

    Public Shared Widening Operator CType(v As Integer) As BaseVal
        Return New BaseVal(v)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim bv As BaseVal = CType(50, BaseVal)
        Console.WriteLine(bv.X)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_ctype_chained_custom_conversions() {
    let src = r#"
Class Meter
    Public Value As Double
    Public Sub New(v As Double)
        Value = v
    End Sub
    Public Shared Widening Operator CType(v As Double) As Meter
        Return New Meter(v)
    End Shared Widening Operator
End Class

Class Kilometer
    Public Value As Double
    Public Sub New(v As Double)
        Value = v
    End Sub
    Public Shared Widening Operator CType(m As Meter) As Kilometer
        Return New Kilometer(m.Value / 1000.0)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim m As Meter = CType(2500.0, Meter)
        Dim km As Kilometer = CType(m, Kilometer)
        Console.WriteLine(km.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2.5"]);
}

#[test]
fn test_vb_ctype_invalid_string_format_throws_format_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim n As Integer = CType("NotANumber", Integer)
        Catch ex As FormatException
            Console.WriteLine("FormatException Caught on CType String")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FormatException Caught on CType String"]);
}

#[test]
fn test_vb_ctype_biginteger_to_decimal_conversion() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim big As BigInteger = 1234567890123456789D
        Dim dec As Decimal = CType(big, Decimal)
        Console.WriteLine(dec.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1234567890123456789"]);
}
