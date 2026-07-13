use super::helpers::run_vb;

#[test]
fn null_conditional_operator_advanced() {
    let out = run_vb(
        r#"
Class Node
    Public Property Value As String
    Public Property NextNode As Node
    Public Function GetName() As String
        Return Value
    End Function
End Class

Module M
    Sub Main()
        Dim root As New Node() With {.Value = "Root"}
        Dim empty As Node = Nothing
        
        ' Null conditional method call
        Console.WriteLine(root?.GetName())
        Console.WriteLine(empty?.GetName() Is Nothing)
        
        ' Null conditional indexing (if it was an array/list)
        Dim arr() As Integer = {1, 2, 3}
        Dim emptyArr() As Integer = Nothing
        
        Console.WriteLine(arr?(0))
        ' We can't really print Nothing directly for integer in VB without it being 0 if not nullable, 
        ' but for arrays ?. indexing returns Nullable(Of T)
        Console.WriteLine(emptyArr?(0).HasValue)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Root", "True", "1", "False"]);
}
