use super::helpers::run_vb;

#[test]
fn structs_nested() {
    let out = run_vb(
        r#"
Structure Outer
    Public X As Integer
    
    Structure Inner
        Public Y As Integer
    End Structure
    
    Public InnerData As Inner
End Structure

Module M
    Sub Main()
        Dim o As New Outer()
        o.X = 10
        o.InnerData.Y = 20
        
        Console.WriteLine(o.X)
        Console.WriteLine(o.InnerData.Y)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn structs_nested_initializers() {
    let out = run_vb(
        r#"
Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Structure Rectangle
    Public TopLeft As Point
    Public BottomRight As Point
End Structure

Module M
    Sub Main()
        Dim r As New Rectangle() With {
            .TopLeft = New Point() With {.X = 0, .Y = 10},
            .BottomRight = New Point() With {.X = 10, .Y = 0}
        }
        Console.WriteLine(r.TopLeft.Y)
        Console.WriteLine(r.BottomRight.X)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "10"]);
}
