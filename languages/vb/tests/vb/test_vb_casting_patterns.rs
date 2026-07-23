use super::helpers::run_vb;

#[test]
fn directcast_downcast_preserves_reference() {
    let out = run_vb(
        r#"
Imports System

Class Animal
End Class

Class Dog
    Inherits Animal
    Public ReadOnly Property Name As String
    Public Sub New(name As String)
        Me.Name = name
    End Sub
End Class

Module M
    Sub Main()
        Dim a As Animal = New Dog("Rex")
        Dim d As Dog = DirectCast(a, Dog)
        Console.WriteLine(d.Name)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["Rex"]);
}

#[test]
fn directcast_incompatible_reference_throws_invalid_cast() {
    let out = run_vb(
        r#"
Imports System

Class A
End Class

Class B
End Class

Module M
    Sub Main()
        Dim a As Object = New A()
        Try
            Dim b As B = DirectCast(a, B)
            Console.WriteLine("NoCast")
        Catch ex As InvalidCastException
            Console.WriteLine("CastFailed")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["CastFailed"]);
}

#[test]
fn ctype_with_boxed_object() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim o As Object = 42
        Console.WriteLine(CType(o, Integer))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["42"]);
}

#[test]
fn ctype_throws_invalid_cast_for_incompatible_value() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim o As Object = True
        Try
            Console.WriteLine(CType(o, Integer))
            Console.WriteLine("NoCast")
        Catch ex As InvalidCastException
            Console.WriteLine("CastFailed")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["CastFailed"]);
}

#[test]
fn trycast_matches_compatible_reference() {
    let out = run_vb(
        r#"
Imports System

Class Base
End Class

Class Derived
    Inherits Base
    Public ReadOnly Property Tag As String = "ok"
End Class

Module M
    Sub Main()
        Dim b As Base = New Derived()
        Dim d As Derived = TryCast(b, Derived)
        Console.WriteLine(d IsNot Nothing)
        Console.WriteLine(d.Tag)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "ok"]);
}

#[test]
fn trycast_returns_nothing_for_incompatible_reference() {
    let out = run_vb(
        r#"
Imports System

Class Base
End Class

Class Other
End Class

Module M
    Sub Main()
        Dim b As Base = New Base()
        Dim d As Other = TryCast(b, Other)
        Console.WriteLine(d Is Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn typeof_identifies_reference_type() {
    let out = run_vb(
        r#"
Imports System

Class Animal
End Class

Module M
    Sub Main()
        Dim o As Object = New Animal()
        Console.WriteLine(TypeOf o Is Animal)
        Console.WriteLine(TypeOf o IsNot Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn typeof_detects_value_type() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim boxed As Object = 99
        Console.WriteLine(TypeOf boxed Is Integer)
        Console.WriteLine(TypeOf boxed Is Decimal)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn direct_vs_try_cast_difference_with_incompatible_types() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim boxed As Object = 99
        Dim castResult As Object = TryCast(boxed, String)
        Console.WriteLine(castResult Is Nothing)

        Try
            Dim direct As String = DirectCast(boxed, String)
            Console.WriteLine(direct)
        Catch ex As InvalidCastException
            Console.WriteLine("DirectFailed")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "DirectFailed"]);
}

#[test]
fn string_and_char_casting_via_ctype() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim o As Object = "x"
        Dim s As String = CType(o, String)
        Dim c As Char = CType(s(0), Char)
        Console.WriteLine(s)
        Console.WriteLine(c)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["x", "x"]);
}
