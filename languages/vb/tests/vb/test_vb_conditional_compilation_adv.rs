use super::helpers::run_vb;

#[test]
fn conditional_compilation_adv() {
    let out = run_vb(
        r#"
#Const DEBUG = True

Module M
    Sub Main()
#If DEBUG Then
        Console.WriteLine("DebugMode")
#ElseIf RELEASE Then
        Console.WriteLine("ReleaseMode")
#Else
        Console.WriteLine("OtherMode")
#End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["DebugMode"]);
}
