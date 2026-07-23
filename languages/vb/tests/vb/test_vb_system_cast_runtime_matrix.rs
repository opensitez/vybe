use super::helpers::run_vb;

#[test]
fn cast_runtime_convert_via_cstr_and_cint() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(CInt("123"))
        Console.WriteLine(CStr(456))
        Console.WriteLine(CBool(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["123", "456", "True"]);
}

#[test]
fn cast_runtime_directcast_downcast_success() {
    let out = run_vb(
        r#"
Class Base
End Class

Class Child
    Inherits Base
    Public Property Name As String = "child"
End Class

Module M
    Sub Main()
        Dim b As Base = New Child()
        Dim c As Child = DirectCast(b, Child)

        Console.WriteLine(c.Name)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["child"]);
}

#[test]
fn cast_runtime_directcast_fails_fast() {
    let out = run_vb(
        r#"
Class A
End Class

Class B
Inherits A
End Class

Module M
    Sub Main()
        Dim b As A = New A()

        Try
            Dim c As B = DirectCast(b, B)
            Console.WriteLine("bad")
        Catch ex As Exception
            Console.WriteLine(TypeOf ex Is InvalidCastException)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn cast_runtime_trycast_nullable_chain() {
    let out = run_vb(
        r#"
Class Leaf
    Inherits Object
End Class

Module M
    Sub Main()
        Dim obj As Object = Nothing
        Dim leaf As Leaf = TryCast(obj, Leaf)
        Console.WriteLine(leaf Is Nothing)

        obj = New Leaf()
        Dim asLeaf As Leaf = TryCast(obj, Leaf)
        Console.WriteLine(asLeaf Is Nothing = False)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn cast_runtime_ctype_with_array_shape() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim boxed As Object = {1, 2, 3}
        Dim values() As Integer = CType(boxed, Integer())
        Console.WriteLine(values.Length)
        Console.WriteLine(values(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2"]);
}
