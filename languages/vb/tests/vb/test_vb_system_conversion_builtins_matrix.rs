use super::helpers::run_vb;

#[test]
fn conversion_builtins_numeric_roundtrip_is_stable() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.ToInt32(123))
        Console.WriteLine(Convert.ToInt64(12.5))
        Console.WriteLine(Convert.ToDecimal("10.25"))
        Console.WriteLine(Convert.ToString(7.5))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["123", "12", "10.25", "7.5"]);
}

#[test]
fn conversion_builtins_boolean_and_binary_parse() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.ToBoolean("true"))
        Console.WriteLine(Convert.ToByte("200"))
        Console.WriteLine(Convert.ToInt16("-12"))
        Console.WriteLine(Convert.ToString(False))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "200", "-12", "False"]);
}

#[test]
fn conversion_builtins_date_and_time_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim d As DateTime = Convert.ToDateTime("2026-07-21T12:00:00")
        Console.WriteLine(d.Year)
        Console.WriteLine(d.Month)
        Console.WriteLine(d.Day)
        Console.WriteLine(Convert.ToString(d.Date))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2026", "7", "21", "07/21/2026"]);
}

#[test]
fn conversion_builtins_base64_roundtrip() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim payload() As Byte = Encoding.UTF8.GetBytes("vb")
        Dim encoded As String = Convert.ToBase64String(payload)
        Dim decoded() As Byte = Convert.FromBase64String(encoded)

        Console.WriteLine(encoded.Length > 0)
        Console.WriteLine(Encoding.UTF8.GetString(decoded))
        Console.WriteLine(Convert.ToBase64String(Encoding.UTF8.GetBytes("")) = "")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "vb", "True"]);
}

#[test]
fn conversion_builtins_to_byte_and_back() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim b As Byte = CByte(255)
        Console.WriteLine(b)
        Console.WriteLine(CStr(b))
        Console.WriteLine(Convert.ToString(CByte("12")))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["255", "255", "12"]);
}

#[test]
fn conversion_builtins_hex_parse_and_to_string() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As Integer = Convert.ToInt32("ff", 16)
        Dim hex As String = Convert.ToString(value, 16)

        Console.WriteLine(value)
        Console.WriteLine(hex)
        Console.WriteLine(hex = "ff")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["255", "ff", "True"]);
}
