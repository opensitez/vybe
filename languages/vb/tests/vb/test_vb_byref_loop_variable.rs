use super::helpers::run_vb;

#[test]
fn byref_loop_variable() {
    let out = run_vb(
        r#"
Module M
    Sub ModifyByRef(ByRef val As Integer)
        val += 10
    End Sub

    Sub Main()
        ' VB.NET allows passing loop variables ByRef, but modifying it inside the method 
        ' behaves exactly like changing the loop variable directly.
        For i As Integer = 1 To 2
            ModifyByRef(i)
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#,
    );
    // Loop variable starts at 1. Modified to 11. Loop prints 11.
    // Next iteration increments to 12. 12 > 2, loop terminates.
    assert_eq!(out, vec!["11"]);
}
