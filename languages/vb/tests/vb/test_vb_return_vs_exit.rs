use super::helpers::run_vb;

#[test]
fn return_vs_exit_function() {
    let out = run_vb(
        r#"
Module M
    Function TestExit() As Integer
        TestExit = 10 ' Implicit return variable
        Exit Function ' Returns immediately
        TestExit = 20
    End Function

    Function TestReturn() As Integer
        Return 30 ' Explicit return
        TestReturn = 40
    End Function

    Function TestImplicit() As Integer
        TestImplicit = 50
        ' Reaches end of function, returns TestImplicit
    End Function

    Sub Main()
        Console.WriteLine(TestExit())
        Console.WriteLine(TestReturn())
        Console.WriteLine(TestImplicit())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "30", "50"]);
}
