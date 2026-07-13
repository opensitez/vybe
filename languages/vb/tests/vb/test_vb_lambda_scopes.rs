use super::helpers::run_vb;

#[test]
fn lambda_closures() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim funcs As New System.Collections.Generic.List(Of Func(Of Integer))
        
        ' VB.NET For loop variable is captured per iteration in modern versions (like C#)
        ' Actually, in VB.NET, the loop variable is declared outside the loop if not explicitly scoped
        ' Let's declare it inside the loop to be safe and test capture semantics.
        For i As Integer = 1 To 3
            Dim captured = i
            funcs.Add(Function() captured * 2)
        Next
        
        For Each f In funcs
            Console.WriteLine(f())
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "4", "6"]);
}

#[test]
fn lambda_closures_loop_variable() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim funcs As New System.Collections.Generic.List(Of Func(Of Integer))
        
        ' In VB.NET, capturing the loop variable directly inside a For loop captures the same variable
        ' so it will evaluate to the final value (4).
        For i As Integer = 1 To 3
            funcs.Add(Function() i * 2)
        Next
        
        For Each f In funcs
            Console.WriteLine(f())
        Next
    End Sub
End Module
"#,
    );
    // After the loop, i is 4 (the value that failed the loop condition `1 To 3`)
    assert_eq!(out, vec!["8", "8", "8"]);
}
