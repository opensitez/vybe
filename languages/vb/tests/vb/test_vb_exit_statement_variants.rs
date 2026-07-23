use super::helpers::run_vb;

#[test]
fn exit_statement_variants() {
    let out = run_vb(
        r#"
Module M
    Sub TestExitSub()
        Console.WriteLine("Sub1")
        Exit Sub
        Console.WriteLine("Sub2")
    End Sub

    Function TestExitFunction() As Integer
        TestExitFunction = 10
        Exit Function
        TestExitFunction = 20
    End Function

    Sub Main()
        TestExitSub()
        Console.WriteLine(TestExitFunction())
        
        For i = 1 To 5
            If i = 3 Then Exit For
            Console.WriteLine("For " & i)
        Next
        
        Dim j = 1
        Do While j <= 5
            If j = 2 Then Exit Do
            Console.WriteLine("Do " & j)
            j += 1
        Loop
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Sub1", "10", "For 1", "For 2", "Do 1"]);
}
