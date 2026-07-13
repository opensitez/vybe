use super::helpers::run_vb;

#[test]
fn addhandler_removehandler_basic() {
    let out = run_vb(
        r#"
Class Button
    Public Event Click()
    
    Public Sub PerformClick()
        RaiseEvent Click()
    End Sub
End Class

Module M
    Sub OnClick1()
        Console.WriteLine("Click 1")
    End Sub
    
    Sub OnClick2()
        Console.WriteLine("Click 2")
    End Sub

    Sub Main()
        Dim btn As New Button()
        
        AddHandler btn.Click, AddressOf OnClick1
        AddHandler btn.Click, AddressOf OnClick2
        btn.PerformClick()
        
        RemoveHandler btn.Click, AddressOf OnClick1
        btn.PerformClick()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Click 1", "Click 2", "Click 2"]);
}
