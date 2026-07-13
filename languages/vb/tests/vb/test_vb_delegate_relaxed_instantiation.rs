use super::helpers::run_vb;

#[test]
fn delegate_relaxed_instantiation() {
    let out = run_vb(
        r#"
Module M
    Delegate Sub ActionDelegate()
    
    Sub MyMethod()
        Console.WriteLine("Invoked")
    End Sub
    
    Sub Main()
        ' Relaxed delegate instantiation (AddressOf automatically creates the delegate type)
        Dim d1 As ActionDelegate = AddressOf MyMethod
        d1.Invoke()
        
        ' Passing to a method expecting a delegate
        ExecuteAction(AddressOf MyMethod)
    End Sub
    
    Sub ExecuteAction(action As ActionDelegate)
        action()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Invoked", "Invoked"]);
}
