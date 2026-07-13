use super::helpers::run_vb;

#[test]
fn multiline_lambdas() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Multiline Sub lambda
        Dim printSum = Sub(x As Integer, y As Integer)
                           Dim result = x + y
                           Console.WriteLine(result)
                       End Sub
                       
        ' Multiline Function lambda
        Dim getGreeting = Function(name As String) As String
                              Dim prefix = "Hello, "
                              Return prefix & name
                          End Function
                          
        printSum(5, 7)
        Console.WriteLine(getGreeting("Alice"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["12", "Hello, Alice"]);
}
