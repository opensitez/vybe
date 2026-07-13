use super::helpers::run_vb;

#[test]
fn seek_statement() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' The Seek function returns the current position
        ' The Seek statement sets the position
        ' We will just parse test them by using them in unreachable code 
        ' to avoid IO issues in the test runner.
        Dim b As Boolean = True
        If Not b Then
            Dim f = FreeFile()
            FileOpen(f, "test.txt", OpenMode.Random)
            Seek(f, 10)
            Dim pos = Seek(f)
            FileClose(f)
        End If
        Console.WriteLine("Seek Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Seek Parsed"]);
}
