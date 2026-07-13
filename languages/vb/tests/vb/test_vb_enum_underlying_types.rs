use super::helpers::run_vb;

#[test]
fn enum_underlying_types() {
    let out = run_vb(
        r#"
' Enum with explicit underlying type Byte
Enum Status As Byte
    Active = 1
    Inactive = 2
    Pending = 3
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Active
        
        ' Check underlying type
        Console.WriteLine(s.GetTypeCode().ToString())
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Byte", "Active"]);
}
