use super::helpers::run_vb;

#[test]
fn catch_when_filters() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim errorCode As Integer = 404
        
        Try
            Throw New System.Exception("HTTP Error")
        Catch ex As System.Exception When errorCode = 500
            Console.WriteLine("Server Error")
        Catch ex As System.Exception When errorCode = 404
            Console.WriteLine("Not Found")
        Catch ex As System.Exception
            Console.WriteLine("Other Error")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Not Found"]);
}
