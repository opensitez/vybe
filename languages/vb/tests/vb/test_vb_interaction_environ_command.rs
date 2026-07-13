use super::helpers::run_vb;

#[test]
fn interaction_environ_command() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Environ gets an environment variable
        Dim pathVar = Environ("PATH")
        Console.WriteLine(pathVar IsNot Nothing)
        
        ' Command gets the command line arguments as a string
        Dim cmd = Command()
        Console.WriteLine(cmd IsNot Nothing)
    End Sub
End Module
"#,
    );
    // Might be empty string, but won't be Nothing
    assert_eq!(out, vec!["True", "True"]);
}
