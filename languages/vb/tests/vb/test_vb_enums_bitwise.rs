use super::helpers::run_vb;

#[test]
fn enums_bitwise_operations() {
    let out = run_vb(
        r#"
<Flags>
Enum Permissions As Byte
    None = 0
    Read = 1
    Write = 2
    Execute = 4
End Enum

Module M
    Sub Main()
        Dim p As Permissions = Permissions.Read Or Permissions.Write
        
        Console.WriteLine(p.HasFlag(Permissions.Read))
        Console.WriteLine(p.HasFlag(Permissions.Execute))
        
        ' Bitwise And to check flag
        Dim isWrite = (p And Permissions.Write) = Permissions.Write
        Console.WriteLine(isWrite)
        
        ' Removing a flag
        p = p And Not Permissions.Read
        Console.WriteLine(p.HasFlag(Permissions.Read))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False", "True", "False"]);
}
