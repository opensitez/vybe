use super::helpers::run_vb;

#[test]
fn trycast_with_interfaces() {
    let out = run_vb(
        r#"
Interface ITest
End Interface

Class A
    Implements ITest
End Class

Class B
End Class

Module M
    Sub Main()
        Dim objA As Object = New A()
        Dim objB As Object = New B()
        
        Dim tA As ITest = TryCast(objA, ITest)
        Console.WriteLine(tA IsNot Nothing)
        
        Dim tB As ITest = TryCast(objB, ITest)
        Console.WriteLine(tB IsNot Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn directcast_with_primitives_should_fail_compilation_but_testing_runtime() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = 42
        
        ' DirectCast requires exact type match for value types boxed in Object
        Dim i As Integer = DirectCast(obj, Integer)
        Console.WriteLine(i)
        
        Try
            ' This throws InvalidCastException at runtime
            Dim d As Double = DirectCast(obj, Double)
            Console.WriteLine(d)
        Catch ex As Exception
            Console.WriteLine("Cast Failed")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "Cast Failed"]);
}
