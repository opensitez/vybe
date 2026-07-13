use super::helpers::run_vb;

#[test]
fn operator_overloading_custom() {
    let out = run_vb(
        r#"
Class Vector
    Public X As Integer
    Public Y As Integer
    
    Public Shared Operator +(v1 As Vector, v2 As Vector) As Vector
        Return New Vector() With {.X = v1.X + v2.X, .Y = v1.Y + v2.Y}
    End Operator
    
    Public Shared Operator -(v1 As Vector) As Vector
        Return New Vector() With {.X = -v1.X, .Y = -v1.Y}
    End Operator
End Class

Module M
    Sub Main()
        Dim v1 As New Vector() With {.X = 1, .Y = 2}
        Dim v2 As New Vector() With {.X = 3, .Y = 4}
        
        Dim v3 = v1 + v2
        Dim v4 = -v1
        
        Console.WriteLine(v3.X)
        Console.WriteLine(v4.X)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4", "-1"]);
}
