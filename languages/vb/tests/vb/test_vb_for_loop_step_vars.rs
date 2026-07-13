use super::helpers::run_vb;

#[test]
fn for_loop_step_decimal() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Using Decimal for exact step
        For i As Decimal = 0D To 1D Step 0.5D
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "0.5", "1"]);
}

#[test]
fn for_loop_step_negative_variable() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim startVal = 5
        Dim endVal = 1
        Dim stepVal = -2
        
        For i As Integer = startVal To endVal Step stepVal
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "3", "1"]);
}

#[test]
fn for_loop_modification_during_loop() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim endVal = 3
        For i As Integer = 1 To endVal
            Console.WriteLine(i)
            endVal = 10 ' modifying endVal has no effect on the loop condition
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
