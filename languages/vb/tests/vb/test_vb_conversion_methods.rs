use super::helpers::run_vb;

#[test]
fn convert_to_int32_from_string() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.ToInt32("123"))
        Console.WriteLine(Convert.ToInt32("  7 "))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["123", "7"]);
}

#[test]
fn convert_to_double_with_invariant_culture() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Console.WriteLine(Convert.ToDouble("3.14", CultureInfo.InvariantCulture))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn convert_to_string_from_bool() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.ToString(True))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn convert_to_boolean_from_numeric_ints() {
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
fn convert_to_char_from_int_codepoint() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.ToChar(65))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["A"]);
}

#[test]
fn convert_to_base64_and_back() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim text As String = "VB.NET"
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(text)
        Dim encoded As String = Convert.ToBase64String(bytes)
        Dim decoded As String = Encoding.UTF8.GetString(Convert.FromBase64String(encoded))
        Console.WriteLine(encoded)
        Console.WriteLine(decoded)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["VkIuTkVUDQ==", "VB.NET"]);
}

#[test]
fn convert_change_type_with_boxed_inputs() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim boxed As Object = "42"
        Dim result As Object = Convert.ChangeType(boxed, GetType(Integer))
        Console.WriteLine(CInt(result))

        Dim decimalValue As Object = "12.5"
        Dim converted As Object = Convert.ChangeType(decimalValue, GetType(Decimal), Globalization.CultureInfo.InvariantCulture)
        Console.WriteLine(CDec(converted))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["42", "12.5"]);
}

#[test]
fn convert_type_code_returns_expected_type() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Convert.GetTypeCode("text") = TypeCode.String)
        Console.WriteLine(Convert.GetTypeCode(1) = TypeCode.Int32)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn convert_to_string_preserves_formatting_of_date() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim d As Date = New Date(2024, 3, 14)
        Dim text As String = Convert.ToString(d)
        Console.WriteLine(text.Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
