use super::helpers::run_vb;

#[test]
fn do_loop_while_until() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
        Do While i < 3
            Console.WriteLine(i)
            i += 1
        Loop
        
        Dim j = 0
        Do Until j = 2
            Console.WriteLine(j)
            j += 1
        Loop
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "1", "2", "0", "1"]);
}

#[test]
fn do_loop_condition_at_bottom() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 10
        Do
            Console.WriteLine(i)
            i += 1
        Loop While i < 5 ' Executes at least once
        
        Dim j = 10
        Do
            Console.WriteLine(j)
            j += 1
        Loop Until j > 5 ' Evaluates true on first check, so stops
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "10"]);
}
