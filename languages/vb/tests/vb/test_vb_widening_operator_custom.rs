use super::helpers::run_vb;

#[test]
fn widening_operator_custom() {
    let out = run_vb(
        r#"
Class BigNum
    Public Value As Integer
    
    ' Implicit conversion from Integer
    Public Shared Widening Operator CType(i As Integer) As BigNum
        Return New BigNum() With {.Value = i}
    End Operator
End Class

Module M
    Sub Main()
        ' Implicitly triggers Widening Operator
        Dim b As BigNum = 42
        Console.WriteLine(b.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
