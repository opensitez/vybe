use super::helpers::run_vb;

#[test]
fn explicit_line_continuations_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Explicit line continuation with underscore
        Dim msg As String = "Hello " & _
                            "World"
        Console.WriteLine(msg)
        
        Dim sum = 1 + _
                  2 + _
                  3
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello World", "6"]);
}

#[test]
fn implicit_line_continuations_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Implicit line continuations (VB 10+)
        Dim sum = 1 +
                  2 +
                  3
        Console.WriteLine(sum)
        
        Dim msg = "A" &
                  "B"
        Console.WriteLine(msg)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["6", "AB"]);
}
