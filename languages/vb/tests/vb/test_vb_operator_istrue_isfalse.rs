use super::helpers::run_vb;

#[test]
fn operator_istrue_isfalse() {
    let out = run_vb(
        r#"
Class TriState
    Public Value As Integer ' 0 = False, 1 = True, -1 = Unknown
    
    Public Shared Operator IsTrue(t As TriState) As Boolean
        Return t.Value = 1
    End Operator
    
    Public Shared Operator IsFalse(t As TriState) As Boolean
        Return t.Value = 0
    End Operator
End Class

Module M
    Sub Main()
        Dim t As New TriState() With {.Value = 1}
        
        ' Relies on IsTrue operator
        If t Then
            Console.WriteLine("True")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
