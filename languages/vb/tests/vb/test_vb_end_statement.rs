use super::helpers::run_vb;

#[test]
fn end_statement() {
    let out = run_vb(
        r#"
Module M
    Sub DoSomething()
        Console.WriteLine("Start")
        End ' Terminates execution immediately
        Console.WriteLine("End") ' Unreachable
    End Sub

    Sub Main()
        DoSomething()
        Console.WriteLine("Main End") ' Unreachable
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Start"]);
}
