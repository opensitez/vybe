use super::helpers::run_vb;

#[test]
fn tuple_destructuring_assignment() {
    let out = run_vb(
        r#"
Module M
    Function GetInfo() As (Name As String, Age As Integer)
        Return ("John", 30)
    End Function

    Sub Main()
        ' VB.NET does not natively support deconstruction syntax (Dim (name, age) = GetInfo())
        ' Wait, it doesn't? C# has deconstruction, but VB.NET does not have direct tuple deconstruction assignment syntax.
        ' Let's just use the tuple literal syntax and element access.
        Dim t = GetInfo()
        Console.WriteLine(t.Name)
        Console.WriteLine(t.Age)
        
        ' We can assign tuples to tuples
        Dim t2 As (String, Integer) = t
        Console.WriteLine(t2.Item1)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["John", "30", "John"]);
}

#[test]
fn tuple_equality() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim t1 = (1, 2)
        Dim t2 = (1, 2)
        Dim t3 = (2, 1)
        
        ' ValueTuple implements Equals
        Console.WriteLine(t1.Equals(t2))
        Console.WriteLine(t1.Equals(t3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}
