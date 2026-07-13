use super::helpers::run_vb;

#[test]
fn addressof_generic() {
    let out = run_vb(
        r#"
Module M
    Sub PrintType(Of T)(val As T)
        Console.WriteLine(val.GetType().Name)
    End Sub

    Sub Main()
        Dim act As Action(Of Integer) = AddressOf PrintType(Of Integer)
        act(42)
        
        Dim act2 As Action(Of String) = AddressOf PrintType(Of String)
        act2("Hello")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Int32", "String"]);
}
