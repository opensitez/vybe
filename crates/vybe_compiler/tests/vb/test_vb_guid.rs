use super::helpers::run_vb;

#[test]
fn vb_system_guid_empty_round_trips_zero_value() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(System.Guid.Empty.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["00000000-0000-0000-0000-000000000000".to_string()]
    );
}

#[test]
fn vb_guid_newguid_uses_shared_dotnet_surface() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text = Guid.NewGuid().ToString()
        Console.WriteLine(text.Length)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["36".to_string()]);
}
