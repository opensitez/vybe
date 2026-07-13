use super::helpers::run_vb;

#[test]
fn operator_overloading_unary() {
    let out = run_vb(
        r#"
Class Vector
    Public X, Y As Integer
    
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    
    ' Unary operator -
    Public Shared Operator -(v As Vector) As Vector
        Return New Vector(-v.X, -v.Y)
    End Operator
    
    ' Unary operator Not
    Public Shared Operator Not(v As Vector) As Vector
        Return New Vector(Not v.X, Not v.Y)
    End Operator
End Class

Module M
    Sub Main()
        Dim v As New Vector(5, -10)
        Dim vNeg = -v
        Console.WriteLine(vNeg.X)
        Console.WriteLine(vNeg.Y)
        
        Dim vNot = Not v
        Console.WriteLine(vNot.X) ' Not 5 = -6
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["-5", "10", "-6"]);
}

#[test]
fn operator_overloading_binary_advanced() {
    let out = run_vb(
        r#"
Class Vector
    Public X, Y As Integer
    
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    
    ' Binary operator * (scalar multiplication)
    Public Shared Operator *(v As Vector, scalar As Integer) As Vector
        Return New Vector(v.X * scalar, v.Y * scalar)
    End Operator
    
    ' Binary operator * (scalar multiplication reversed)
    Public Shared Operator *(scalar As Integer, v As Vector) As Vector
        Return New Vector(v.X * scalar, v.Y * scalar)
    End Operator
End Class

Module M
    Sub Main()
        Dim v As New Vector(2, 3)
        Dim v1 = v * 5
        Dim v2 = 10 * v
        
        Console.WriteLine(v1.X)
        Console.WriteLine(v2.Y)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "30"]);
}
