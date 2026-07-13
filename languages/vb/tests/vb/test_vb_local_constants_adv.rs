use super::helpers::run_vb;

#[test]
fn local_constants_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Const inside a method
        Const MaxLimit As Integer = 100
        Const Greeting As String = "Hello"
        
        Console.WriteLine(MaxLimit)
        Console.WriteLine(Greeting)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100", "Hello"]);
}
