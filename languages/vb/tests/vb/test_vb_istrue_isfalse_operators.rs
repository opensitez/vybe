use super::helpers::run_vb;

#[test]
fn istrue_isfalse_operators() {
    let out = run_vb(
        r#"
Class Truthy
    Public Value As Integer
    
    Public Shared Operator IsTrue(ByVal obj As Truthy) As Boolean
        Return obj.Value > 0
    End Operator
    
    Public Shared Operator IsFalse(ByVal obj As Truthy) As Boolean
        Return obj.Value <= 0
    End Operator
End Class

Module M
    Sub Main()
        Dim t1 As New Truthy() With {.Value = 10}
        Dim t2 As New Truthy() With {.Value = -5}
        
        If t1 Then
            Console.WriteLine("t1 is true")
        End If
        
        If Not t2 Then
            Console.WriteLine("t2 is false")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["t1 is true", "t2 is false"]);
}
