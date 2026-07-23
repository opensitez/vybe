use super::helpers::run_vb;

#[test]
fn guid_newguid_is_not_empty() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim g As Guid = Guid.NewGuid()
        Console.WriteLine(g = Guid.Empty)
        Console.WriteLine(g.ToString("N").Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "32"]);
}

#[test]
fn guid_parse_valid_string_roundtrips() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "d87a74a4-5694-4d8b-a3ed-3085794711f1"
        Dim g As Guid = Guid.Parse(text)
        Console.WriteLine(g.ToString("D").ToLower())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["d87a74a4-5694-4d8b-a3ed-3085794711f1"]);
}

#[test]
fn guid_parse_with_braces_is_accepted() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim g As Guid = Guid.Parse("{d87a74a4-5694-4d8b-a3ed-3085794711f1}")
        Console.WriteLine(g.ToString("N"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["d87a74a4-5694-4d8b-a3ed-3085794711f1"]);
}

#[test]
fn guid_try_parse_valid_and_invalid() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim g As Guid
        Console.WriteLine(Guid.TryParse("not-a-guid", g))
        Console.WriteLine(Guid.TryParse("d87a74a4-5694-4d8b-a3ed-3085794711f1", g))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn guid_newguid_is_unique_across_calls() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim first As Guid = Guid.NewGuid()
        Dim second As Guid = Guid.NewGuid()
        Console.WriteLine(first = second)
        Console.WriteLine(first <> Guid.Empty)
        Console.WriteLine(second <> Guid.Empty)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True", "True"]);
}
