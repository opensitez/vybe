use super::helpers::run_vb;

#[test]
fn delegate_creation_addressof() {
    let out = run_vb(
        r#"
Delegate Function MathOp(x As Integer, y As Integer) As Integer

Module M
    Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function

    Sub Main()
        Dim op As MathOp = AddressOf Add
        Console.WriteLine(op(5, 3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn delegate_creation_lambda() {
    let out = run_vb(
        r#"
Delegate Function MathOp(x As Integer, y As Integer) As Integer

Module M
    Sub Main()
        Dim op As MathOp = Function(a, b) a * b
        Console.WriteLine(op(5, 3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn delegate_multicast() {
    let out = run_vb(
        r#"
Delegate Sub Log(msg As String)

Module M
    Sub Log1(msg As String)
        Console.WriteLine("1: " & msg)
    End Sub
    
    Sub Log2(msg As String)
        Console.WriteLine("2: " & msg)
    End Sub

    Sub Main()
        Dim logger As Log = AddressOf Log1
        logger = CType([Delegate].Combine(logger, New Log(AddressOf Log2)), Log)
        logger("Test")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1: Test", "2: Test"]);
}
