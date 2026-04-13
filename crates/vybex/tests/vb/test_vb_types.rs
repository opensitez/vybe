use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Types — Enum, Structure, Interface, casting
// ═══════════════════════════════════════════════════════════

#[test]
fn enum_basic() {
    let out = run_vb(r#"
Enum Color
    Red
    Green
    Blue
End Enum

Module M
    Sub Main()
        Dim c As Color = Color.Green
        Console.WriteLine(c)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_with_values() {
    let out = run_vb(r#"
Enum Status
    Active = 1
    Inactive = 0
    Pending = 2
End Enum

Module M
    Sub Main()
        Dim s As Status = Status.Pending
        Console.WriteLine(s)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn structure_basic() {
    let out = run_vb(r#"
Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Module M
    Sub Main()
        Dim p As New Point()
        p.X = 10
        p.Y = 20
        Console.WriteLine(p.X + p.Y)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn interface_basic() {
    let out = run_vb(r#"
Interface IGreeter
    Function Greet() As String
End Interface

Class HelloGreeter
    Implements IGreeter
    Public Function Greet() As String
        Return "Hello!"
    End Function
End Class

Module M
    Sub Main()
        Dim g As New HelloGreeter()
        Console.WriteLine(g.Greet())
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello!"]);
}

#[test]
fn cint_conversion() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim s As String = "42"
        Dim n As Integer = CInt(s)
        Console.WriteLine(n + 8)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn cstr_conversion() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim n As Integer = 100
        Dim s As String = CStr(n)
        Console.WriteLine(s & " items")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["100 items"]);
}

#[test]
fn isnothing_check() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim s As String = Nothing
        Console.WriteLine(IsNothing(s))
        s = "hello"
        Console.WriteLine(IsNothing(s))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn type_integer_operations() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim a As Integer = 10
        Dim b As Integer = 3
        Console.WriteLine(a + b)
        Console.WriteLine(a - b)
        Console.WriteLine(a * b)
        Console.WriteLine(a \ b)
        Console.WriteLine(a Mod b)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["13", "7", "30", "3", "1"]);
}

#[test]
fn type_double_operations() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim a As Double = 10.5
        Dim b As Double = 3.2
        Console.WriteLine(a + b)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["13.7"]);
}

#[test]
fn type_boolean() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim t As Boolean = True
        Dim f As Boolean = False
        Console.WriteLine(t)
        Console.WriteLine(f)
        Console.WriteLine(t And f)
        Console.WriteLine(t Or f)
        Console.WriteLine(Not t)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["true", "false", "false", "true", "false"]);
}

#[test]
fn constant_declaration() {
    let out = run_vb(r#"
Module M
    Const PI As Double = 3.14159
    Sub Main()
        Console.WriteLine(PI)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn nothing_assignment() {
    let out = run_vb(r#"
Class MyClass
    Public Value As Integer = 42
End Class

Module M
    Sub Main()
        Dim obj As New MyClass()
        Console.WriteLine(obj.Value)
        obj = Nothing
        Console.WriteLine(IsNothing(obj))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["42", "true"]);
}
