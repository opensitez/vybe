use super::helpers::run_vb;

#[test]
fn my_app_log() {
    let out = run_vb(
        r#"
Module M
    ' Simulate My.Application.Log
    Class LogClass
        Public Sub WriteEntry(msg As String)
            Console.WriteLine("Log: " & msg)
        End Sub
    End Class
    
    Class AppClass
        Public Log As New LogClass()
    End Class
    
    Dim Application As New AppClass()
    
    Sub Main()
        Application.Log.WriteEntry("Started")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Log: Started"]);
}
