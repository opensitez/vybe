use super::helpers::run_vb;

#[test]
fn module_const_enums() {
    let out = run_vb(
        r#"
Module Constants
    Public Const Pi As Double = 3.14159
    Public Const AppName As String = "MyApp"
    
    Public Enum Mode
        Fast = 1
        Safe = 2
    End Enum
End Module

Module M
    Sub Main()
        ' Accessing module members directly
        Console.WriteLine(Pi)
        Console.WriteLine(AppName)
        
        Dim m As Mode = Mode.Fast
        Console.WriteLine(m)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3.14159", "MyApp", "Fast"]);
}
