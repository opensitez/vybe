use super::helpers::run_vb;

#[test]
fn narrowing_operators() {
    let out = run_vb(
        r#"
Class Wrapper
    Public Value As Integer
    
    ' Narrowing explicitly requires casting
    Public Shared Narrowing Operator CType(w As Wrapper) As Integer
        Return w.Value
    End Operator
    
    ' Widening allows implicit casting
    Public Shared Widening Operator CType(i As Integer) As Wrapper
        Return New Wrapper() With {.Value = i}
    End Operator
End Class

Module M
    Sub Main()
        ' Widening implicit
        Dim w As Wrapper = 42
        Console.WriteLine(w.Value)
        
        ' Narrowing explicit
        Dim i As Integer = CType(w, Integer)
        Console.WriteLine(i)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "42"]);
}
