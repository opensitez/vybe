use super::helpers::run_vb;

#[test]
fn extension_method_string_transforms_are_visible() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices

Module TextExtensions
    <Extension()>
    Public Function Wrap(value As String, prefix As String, suffix As String) As String
        Return prefix & value & suffix
    End Function

    <Extension()>
    Public Function ReverseText(value As String) As String
        Dim chars As Char() = value.ToCharArray()
        Array.Reverse(chars)
        Return New String(chars)
    End Function
End Module

Module M
    Sub Main()
        Dim original As String = "abc"
        Console.WriteLine(original.Wrap("[", "]"))
        Console.WriteLine(original.ReverseText())
    End Module
End Module
"#,
    );

    assert_eq!(out, vec!["[abc]", "cba"]);
}

#[test]
fn extension_method_integer_square() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices

Module MathExtensions
    <Extension()>
    Public Function Square(value As Integer) As Integer
        Return value * value
    End Function

    <Extension()>
    Public Function DoubleValue(value As Integer) As Integer
        Return value + value
    End Function
End Module

Module M
    Sub Main()
        Dim n As Integer = 7
        Console.WriteLine(n.Square())
        Console.WriteLine(n.DoubleValue())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["49", "14"]);
}

#[test]
fn extension_method_generic_like_list_projection() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices

Module EnumerableExtensions
    <Extension()>
    Public Function FirstOrDefaultIfExists(values As Integer(), fallback As Integer) As Integer
        If values Is Nothing OrElse values.Length = 0 Then
            Return fallback
        End If
        Return values(0)
    End Function
End Module

Module M
    Sub Main()
        Dim values As Integer() = {12, 24, 36}
        Dim empty As Integer() = {}

        Console.WriteLine(values.FirstOrDefaultIfExists(99))
        Console.WriteLine(empty.FirstOrDefaultIfExists(77))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["12", "77"]);
}
