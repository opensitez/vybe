use super::helpers::run_vb;

#[test]
fn anonymous_delegates() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Anonymous Sub delegate
        Dim log = Sub(msg As String) Console.WriteLine("Log: " & msg)
        
        ' Anonymous Function delegate
        Dim multiply = Function(x As Integer, y As Integer) As Integer
                           Return x * y
                       End Function
        
        log("Test")
        Console.WriteLine(multiply(3, 4))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Log: Test", "12"]);
}
