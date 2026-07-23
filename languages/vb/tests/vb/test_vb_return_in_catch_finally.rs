use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Return inside Catch and Finally
// ═══════════════════════════════════════════════════════════

#[test]
fn return_in_catch_finally() {
    let out = run_vb(
        r#"
Module M
    Function TestReturn() As Integer
        Try
            Throw New Exception("Error")
        Catch ex As Exception
            Return 1
        Finally
            ' VB.NET allows Return in Finally?
            ' Wait, Return in Finally is a compiler error in VB.NET (and C#)!
            ' So we just test Return in Catch and modifying the return value implicitly by assigning to the function name
            Console.WriteLine("Finally executed")
        End Try
    End Function

    Function TestImplicitReturn() As Integer
        Try
            Throw New Exception("Error")
        Catch ex As Exception
            TestImplicitReturn = 2
            Exit Function
        Finally
            Console.WriteLine("Finally executed 2")
        End Try
    End Function

    Sub Main()
        Console.WriteLine(TestReturn())
        Console.WriteLine(TestImplicitReturn())
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["Finally executed", "1", "Finally executed 2", "2"]
    );
}
