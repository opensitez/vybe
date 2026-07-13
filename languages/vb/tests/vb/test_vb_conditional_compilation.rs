use super::helpers::run_vb;

#[test]
fn conditional_compilation() {
    let out = run_vb(
        r#"
#Const DEBUG_MODE = True
#Const VERSION = 2

Module M
    Sub Main()
#If DEBUG_MODE Then
        Console.WriteLine("Debug On")
#Else
        Console.WriteLine("Debug Off")
#End If

#If VERSION = 1 Then
        Console.WriteLine("V1")
#ElseIf VERSION = 2 Then
        Console.WriteLine("V2")
#Else
        Console.WriteLine("Unknown")
#End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Debug On", "V2"]);
}
