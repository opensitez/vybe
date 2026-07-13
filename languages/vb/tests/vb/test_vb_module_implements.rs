use super::helpers::run_vb;

#[test]
fn module_implements() {
    let out = run_vb(
        r#"
Interface IRunnable
    Sub Run()
End Interface

' Modules cannot implement interfaces in standard VB.NET.
' We wrap this in a class to test the syntax for Implements instead.
Class Runner
    Implements IRunnable
    
    Public Sub Run() Implements IRunnable.Run
        Console.WriteLine("Running")
    End Sub
End Class

Module M
    Sub Main()
        Dim r As IRunnable = New Runner()
        r.Run()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Running"]);
}
