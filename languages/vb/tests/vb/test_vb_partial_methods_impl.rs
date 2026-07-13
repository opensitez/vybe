use super::helpers::run_vb;

#[test]
fn partial_methods_impl() {
    let out = run_vb(
        r#"
Partial Class Helper
    Partial Private Sub Log(msg As String)
    End Sub
    
    Public Sub DoWork()
        Console.WriteLine("Working")
        Log("Done")
    End Sub
End Class

Partial Class Helper
    ' Implementation of the partial method
    Private Sub Log(msg As String)
        Console.WriteLine("Log: " & msg)
    End Sub
End Class

Module M
    Sub Main()
        Dim h As New Helper()
        h.DoWork()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Working", "Log: Done"]);
}

#[test]
fn partial_methods_unimplemented() {
    let out = run_vb(
        r#"
Partial Class Helper
    Partial Private Sub Log(msg As String)
    End Sub
    
    Public Sub DoWork()
        Console.WriteLine("Start")
        ' This call is compiled away if not implemented
        Log("Middle")
        Console.WriteLine("End")
    End Sub
End Class

Module M
    Sub Main()
        Dim h As New Helper()
        h.DoWork()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Start", "End"]);
}
