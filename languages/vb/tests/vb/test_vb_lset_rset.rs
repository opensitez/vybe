use super::helpers::run_vb;

#[test]
fn lset_rset_statements() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' LSet and RSet pad strings with spaces to match the length of the target variable
        Dim s1 As String = "1234567890"
        LSet s1 = "Left"
        Console.WriteLine("[" & s1 & "]")
        
        Dim s2 As String = "1234567890"
        RSet s2 = "Right"
        Console.WriteLine("[" & s2 & "]")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["[Left      ]", "[     Right]"]);
}
