use super::helpers::run_vb;

#[test]
fn guid_newguid_is_not_empty() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim g As Guid = Guid.NewGuid()
        Console.WriteLine(g <> Guid.Empty)
        Console.WriteLine(g.ToString("N").Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "32"]);
}

#[test]
fn guid_parse_and_to_string_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim parsed As Guid = Guid.Parse("4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4")
        Dim text As String = parsed.ToString("N")
        Console.WriteLine(text.Length)
        Console.WriteLine(Guid.Parse(text).ToString("D"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["32", "4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4"]);
}

#[test]
fn guid_try_parse_handles_invalid_input() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim parsed As Guid = Guid.Empty
        Dim ok As Boolean = Guid.TryParse("not-a-guid", parsed)
        Console.WriteLine(ok)
        Console.WriteLine(parsed = Guid.Empty)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn guid_from_byte_array_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source As Byte() = {
            &H4F, &H7F, &H1D, &HCB, &H7A, &H39, &H4F, &H9B,
            &HB0, &HDE, &HF7, &HF9, &HA2, &HF5, &HF8, &HF4
        }
        Dim g As New Guid(source)
        Dim bytes() As Byte = g.ToByteArray()
        Dim restored As New Guid(bytes)
        Console.WriteLine(g = restored)
        Console.WriteLine(bytes.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "16"]);
}

#[test]
fn guid_equals_and_hash_contract() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As Guid = Guid.Parse("4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4")
        Dim b As Guid = Guid.Parse("4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4")
        Dim c As Guid = Guid.Parse("8f4db1ef-2f1c-48ab-8b45-ff31a8a8cbe3")
        Console.WriteLine(a = b)
        Console.WriteLine(a <> c)
        Console.WriteLine(a.GetHashCode() <> c.GetHashCode())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn guid_empty_constant_is_not_unique() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Guid.Empty.ToString("D"))
        Console.WriteLine(Guid.Empty = New Guid())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["00000000-0000-0000-0000-000000000000", "True"]);
}

#[test]
fn guid_brace_and_parentheses_forms_parse() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As Guid = Guid.Parse("{4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4}")
        Dim b As Guid = Guid.Parse("(4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4)")
        Console.WriteLine(a = b)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
