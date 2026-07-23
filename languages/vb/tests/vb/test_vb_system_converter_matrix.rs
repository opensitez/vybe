use super::helpers::run_vb;

#[test]
fn converter_to_int32_from_string() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As Integer = Convert.ToInt32("42")
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["42"]);
}

#[test]
fn converter_to_boolean_from_int() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.ToBoolean(1))
        Console.WriteLine(Convert.ToBoolean(0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn converter_decimal_from_string() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As Decimal = Convert.ToDecimal("125")
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["125"]);
}

#[test]
fn converter_to_base64_roundtrip() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim bytes() As Byte = Encoding.UTF8.GetBytes("vb")
        Dim text As String = Convert.ToBase64String(bytes)
        Dim restored As Byte() = Convert.FromBase64String(text)
        Console.WriteLine(restored.Length)
        Console.WriteLine(Encoding.UTF8.GetString(restored))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "vb"]);
}

#[test]
fn converter_change_type_from_string() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As Object = Convert.ChangeType("17", GetType(Integer))
        Dim asInt As Integer = CType(value, Integer)
        Console.WriteLine(asInt)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["17"]);
}

#[test]
fn converter_to_int64_roundtrip_string_and_bytes() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As Int64 = Convert.ToInt64("9223372036854775807")
        Console.WriteLine(value > 0)
        Console.WriteLine(Convert.ToString(value).Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "19"]);
}

#[test]
fn converter_isdbnull() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.IsDBNull(DBNull.Value))
        Console.WriteLine(Convert.IsDBNull("value"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn converter_to_datetime() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim dt As DateTime = Convert.ToDateTime("2026-07-21")
        Console.WriteLine(dt.Year)
        Console.WriteLine(dt.Month)
        Console.WriteLine(dt.Day)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2026", "7", "21"]);
}

#[test]
fn converter_to_string_of_numeric() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim s As String = Convert.ToString(123)
        Console.WriteLine(s)
        Console.WriteLine(s.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["123", "3"]);
}

#[test]
fn converter_to_single_and_double() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As Single = Convert.ToSingle("125")
        Dim b As Double = Convert.ToDouble("250")
        Console.WriteLine(a + b)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["375"]);
}
