use super::helpers::run_vb;

#[test]
fn try_catch_finally_returns() {
    let out = run_vb(
        r#"
Module M
    Function TestReturn() As Integer
        Try
            Return 1
        Catch
            Return 2
        Finally
            ' Cannot Return from Finally in VB.NET, but can write to console
            Console.WriteLine("Finally")
        End Try
    End Function

    Function TestThrow() As Integer
        Try
            Throw New Exception("Error")
        Catch
            Return 3
        Finally
            Console.WriteLine("Finally2")
        End Try
    End Function

    Sub Main()
        Console.WriteLine(TestReturn())
        Console.WriteLine(TestThrow())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Finally", "1", "Finally2", "3"]);
}
