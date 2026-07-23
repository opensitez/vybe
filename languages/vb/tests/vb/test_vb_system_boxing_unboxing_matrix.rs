use super::helpers::run_vb;

#[test]
fn boxing_unboxing_identity_of_integers() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim boxed As Object = 21
        Dim value As Integer = CType(boxed, Integer)

        Console.WriteLine(value)
        Console.WriteLine(TypeOf boxed Is Integer)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["21", "True"]);
}

#[test]
fn boxing_unboxing_reference_preservation_on_classes() {
    let out = run_vb(
        r#"
Class Holder
    Public Value As Integer

    Public Sub New(v As Integer)
        Value = v
    End Sub
End Class

Module M
    Sub Main()
        Dim h As New Holder(9)
        Dim boxed As Object = h
        Dim unboxed As Holder = CType(boxed, Holder)

        Console.WriteLine(unboxed.Value)
        unboxed.Value = 10
        Console.WriteLine(h.Value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["9", "10"]);
}

#[test]
fn boxing_unboxing_trycast_returns_nothing_on_mismatch() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim boxed As Object = 12.5
        Dim intRef As Nullable(Of Integer)

        intRef = TryCast(boxed, Integer)
        Console.WriteLine(intRef.HasValue)

        Dim asObj As String = TryCast(boxed, String)
        Console.WriteLine(asObj Is Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn boxing_unboxing_directcast_throws_when_invalid() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim boxed As Object = "hello"

        Try
            Dim x As Integer = DirectCast(boxed, Integer)
            Console.WriteLine("no")
        Catch ex As InvalidCastException
            Console.WriteLine("invalid")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["invalid"]);
}

#[test]
fn boxing_unboxing_nested_struct_like_record() {
    let out = run_vb(
        r#"
Structure Pnt
    Public X As Integer
    Public Y As Integer
End Structure

Module M
    Sub Main()
        Dim p As Pnt
        p.X = 2
        p.Y = 3

        Dim boxed As Object = p
        Dim restored As Pnt = CType(boxed, Pnt)

        Console.WriteLine(restored.X)
        Console.WriteLine(restored.Y)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn boxing_unboxing_array_of_objects_avoids_aliasing() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim objects() As Object = {1, 2, 3}
        objects(0) = CInt(objects(0)) + 1

        Console.WriteLine(CInt(objects(0)))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}
