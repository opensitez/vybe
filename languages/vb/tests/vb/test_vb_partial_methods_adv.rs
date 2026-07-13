use super::helpers::run_vb;

#[test]
fn partial_methods_adv() {
    let out = run_vb(
        r#"
Partial Class Processor
    ' Declaration of a partial method
    Partial Private Sub Log(msg As String)
    End Sub
    
    Public Sub Run()
        Log("Running")
        Console.WriteLine("Done")
    End Sub
End Class

Partial Class Processor
    ' Implementation of the partial method
    Private Sub Log(msg As String)
        Console.WriteLine("Log: " & msg)
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Processor()
        p.Run()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Log: Running", "Done"]);
}
