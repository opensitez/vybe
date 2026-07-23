use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Convert.ChangeType & Dynamic Type Conversion
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_convert_change_type_string_to_integer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj As Object = "123"
        Dim res As Object = Convert.ChangeType(obj, GetType(Integer))
        Console.WriteLine(res.GetType().Name & ":" & res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32:123"]);
}

#[test]
fn test_vb_convert_change_type_integer_to_double() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj As Object = 500
        Dim res As Object = Convert.ChangeType(obj, GetType(Double))
        Console.WriteLine(res.GetType().Name & ":" & res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Double:500"]);
}

#[test]
fn test_vb_convert_change_type_string_to_date_time() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dateStr As Object = "2025-12-31"
        Dim res As Object = Convert.ChangeType(dateStr, GetType(DateTime), CultureInfo.InvariantCulture)
        Dim dt As DateTime = CType(res, DateTime)
        Console.WriteLine(dt.Year & "-" & dt.Month & "-" & dt.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-12-31"]);
}

#[test]
fn test_vb_convert_change_type_string_to_enum() {
    let src = r#"
Imports System

Enum Priority
    Low = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim val As Object = "High"
        ' Enum.Parse works with ChangeType or Enum.Parse
        Dim res As Object = [Enum].Parse(GetType(Priority), CStr(val))
        Console.WriteLine(res.GetType().Name & ":" & res.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Priority:High"]);
}

#[test]
fn test_vb_convert_change_type_null_returns_null_for_reference_type() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim res As Object = Convert.ChangeType(Nothing, GetType(String))
        Console.WriteLine(res Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_convert_change_type_null_throws_invalid_cast_for_value_type() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Convert.ChangeType(Nothing, GetType(Integer))
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on Null to Int ValueType")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on Null to Int ValueType"]
    );
}

#[test]
fn test_vb_convert_change_type_iconvertible_implementation() {
    let src = r#"
Imports System
Imports System.Globalization

Class CustomConvertible
    Implements IConvertible

    Public Value As Integer = 777

    Public Function ToInt32(provider As IFormatProvider) As Integer Implements IConvertible.ToInt32
        Return Value
    End Function

    Public Function GetTypeCode() As TypeCode Implements IConvertible.GetTypeCode
        Return TypeCode.Object
    End Function
    Public Function ToBoolean(provider As IFormatProvider) As Boolean Implements IConvertible.ToBoolean
        Return Value <> 0
    End Function
    Public Function ToByte(provider As IFormatProvider) As Byte Implements IConvertible.ToByte
        Return CByte(Value)
    End Function
    Public Function ToChar(provider As IFormatProvider) As Char Implements IConvertible.ToChar
        Return "X"c
    End Function
    Public Function ToDateTime(provider As IFormatProvider) As DateTime Implements IConvertible.ToDateTime
        Return DateTime.MinValue
    End Function
    Public Function ToDecimal(provider As IFormatProvider) As Decimal Implements IConvertible.ToDecimal
        Return Value
    End Function
    Public Function ToDouble(provider As IFormatProvider) As Double Implements IConvertible.ToDouble
        Return Value
    End Function
    Public Function ToInt16(provider As IFormatProvider) As Short Implements IConvertible.ToInt16
        Return CShort(Value)
    End Function
    Public Function ToInt64(provider As IFormatProvider) As Long Implements IConvertible.ToInt64
        Return Value
    End Function
    Public Function ToSByte(provider As IFormatProvider) As SByte Implements IConvertible.ToSByte
        Return CSByte(Value)
    End Function
    Public Function ToSingle(provider As IFormatProvider) As Single Implements IConvertible.ToSingle
        Return Value
    End Function
    Public Function ToString(provider As IFormatProvider) As String Implements IConvertible.ToString
        Return Value.ToString()
    End Function
    Public Function ToType(conversionType As Type, provider As IFormatProvider) As Object Implements IConvertible.ToType
        If conversionType Is GetType(Integer) Then Return Value
        Throw New InvalidCastException()
    End Function
    Public Function ToUInt16(provider As IFormatProvider) As UShort Implements IConvertible.ToUInt16
        Return CUShort(Value)
    End Function
    Public Function ToUInt32(provider As IFormatProvider) As UInteger Implements IConvertible.ToUInt32
        Return CUInt(Value)
    End Function
    Public Function ToUInt64(provider As IFormatProvider) As ULong Implements IConvertible.ToUInt64
        Return CULng(Value)
    End Function
End Class

Module Program
    Sub Main()
        Dim cc As New CustomConvertible()
        Dim num As Object = Convert.ChangeType(cc, GetType(Integer))
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["777"]);
}

#[test]
fn test_vb_convert_change_type_culture_info_decimal_point() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim strVal As Object = "12,34"
        Dim deCulture As New CultureInfo("de-DE")
        Dim res As Object = Convert.ChangeType(strVal, GetType(Double), deCulture)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12.34"]);
}

#[test]
fn test_vb_convert_change_type_nullable_underlying_type() {
    let src = r#"
Imports System

Module Program
    Private Function ChangeTypeToNullable(val As Object, targetType As Type) As Object
        Dim underlying = Nullable.GetUnderlyingType(targetType)
        Dim effectiveType = If(underlying, targetType)
        If val Is Nothing Then Return Nothing
        Return Convert.ChangeType(val, effectiveType)
    End Function

    Sub Main()
        Dim res = ChangeTypeToNullable("99", GetType(Integer?))
        Console.WriteLine(res.GetType().Name & ":" & res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32:99"]);
}

#[test]
fn test_vb_convert_change_type_boolean_from_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim strTrue As Object = "True"
        Dim strFalse As Object = "False"
        Dim b1 As Object = Convert.ChangeType(strTrue, GetType(Boolean))
        Dim b2 As Object = Convert.ChangeType(strFalse, GetType(Boolean))
        Console.WriteLine(b1 & "|" & b2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_convert_change_type_invalid_format_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Convert.ChangeType("InvalidNumber", GetType(Integer))
        Catch ex As FormatException
            Console.WriteLine("FormatException Caught on ChangeType")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FormatException Caught on ChangeType"]);
}

#[test]
fn test_vb_convert_change_type_overflow_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Convert.ChangeType(999999, GetType(Byte))
        Catch ex As OverflowException
            Console.WriteLine("OverflowException Caught on ChangeType Byte Overflow")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["OverflowException Caught on ChangeType Byte Overflow"]
    );
}

#[test]
fn test_vb_convert_change_type_guid_parsing() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim guidStr = "d3b07384-d113-43a4-a719-000000000000"
        Dim g As Guid = Guid.Parse(guidStr)
        Console.WriteLine(g.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["d3b07384-d113-43a4-a719-000000000000"]);
}

#[test]
fn test_vb_convert_change_type_time_span_from_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim tsStr = "01:30:00"
        Dim ts As TimeSpan = TimeSpan.Parse(tsStr)
        Console.WriteLine(ts.TotalMinutes)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["90"]);
}

#[test]
fn test_vb_convert_change_type_type_code_enumeration() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim code = Type.GetTypeCode(GetType(Double))
        Console.WriteLine(code.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Double"]);
}

#[test]
fn test_vb_convert_change_type_same_type_returns_same_instance() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim orig As Object = "SameString"
        Dim converted As Object = Convert.ChangeType(orig, GetType(String))
        Console.WriteLine(Object.ReferenceEquals(orig, converted))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_convert_change_type_decimal_to_integer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dec As Object = 88.5D
        Dim num As Object = Convert.ChangeType(dec, GetType(Integer))
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["88"]);
}

#[test]
fn test_vb_convert_change_type_uri_from_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim urlStr = "https://example.com/api"
        Dim u As New Uri(urlStr)
        Console.WriteLine(u.Host & "|" & u.Scheme)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["example.com|https"]);
}

#[test]
fn test_vb_convert_change_type_byte_array_to_string_disallowed() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Object = New Byte() {65, 66}
        Try
            ' Direct ChangeType from Byte() to String throws InvalidCastException!
            Convert.ChangeType(bytes, GetType(String))
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on Byte Array to String")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on Byte Array to String"]
    );
}

#[test]
fn test_vb_convert_change_type_version_parse() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim vStr = "2.5.10.0"
        Dim v As Version = Version.Parse(vStr)
        Console.WriteLine(v.Major & "." & v.Minor)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2.5"]);
}
