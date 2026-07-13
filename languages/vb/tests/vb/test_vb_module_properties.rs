use super::helpers::run_vb;

#[test]
fn module_properties() {
    let out = run_vb(
        r#"
Module GlobalState
    Private _val As Integer = 10
    
    ' Properties in a module are implicitly Shared (static)
    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(val As Integer)
            _val = val
        End Set
    End Property
End Module

Module M
    Sub Main()
        Console.WriteLine(GlobalState.Value)
        GlobalState.Value = 20
        Console.WriteLine(GlobalState.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}
