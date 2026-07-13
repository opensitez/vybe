use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Operator Overloading
// ═══════════════════════════════════════════════════════════

#[test]
fn operator_overloading() {
    let out = run_vb(
        r#"
Structure Vector
    Public X As Integer
    Public Y As Integer
    
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    
    Public Shared Operator +(v1 As Vector, v2 As Vector) As Vector
        Return New Vector(v1.X + v2.X, v1.Y + v2.Y)
    End Operator
    
    Public Shared Operator =(v1 As Vector, v2 As Vector) As Boolean
        Return v1.X = v2.X AndAlso v1.Y = v2.Y
    End Operator
    
    Public Shared Operator <>(v1 As Vector, v2 As Vector) As Boolean
        Return Not (v1 = v2)
    End Operator
End Structure

Module M
    Sub Main()
        Dim v1 As New Vector(1, 2)
        Dim v2 As New Vector(3, 4)
        Dim v3 = v1 + v2
        
        Console.WriteLine(v3.X)
        Console.WriteLine(v3.Y)
        Console.WriteLine(v3 = New Vector(4, 6))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4", "6", "True"]);
}
