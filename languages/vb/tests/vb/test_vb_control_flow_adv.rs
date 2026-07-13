use super::helpers::run_vb;

#[test]
fn control_flow_goto_labels() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
    StartLabel:
        If i = 3 Then
            GoTo EndLabel
        End If
        Console.WriteLine(i)
        i += 1
        GoTo StartLabel
        
    EndLabel:
        Console.WriteLine("Done")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "1", "2", "Done"]);
}

#[test]
fn control_flow_exit_try() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Console.WriteLine("Start")
            Exit Try
            Console.WriteLine("Middle")
        Catch
        Finally
            Console.WriteLine("Finally")
        End Try
        Console.WriteLine("End")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Start", "Finally", "End"]);
}

#[test]
fn control_flow_continue_do_for() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        For i As Integer = 1 To 3
            If i = 2 Then Continue For
            Console.WriteLine("For " & i)
        Next
        
        Dim j = 0
        Do While j < 3
            j += 1
            If j = 2 Then Continue Do
            Console.WriteLine("Do " & j)
        Loop
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["For 1", "For 3", "Do 1", "Do 3"]);
}
