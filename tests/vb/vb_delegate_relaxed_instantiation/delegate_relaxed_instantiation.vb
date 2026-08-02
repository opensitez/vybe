' vybe-test: vb/vb_delegate_relaxed_instantiation/delegate_relaxed_instantiation
' origin: languages/vb/tests/vb/test_vb_delegate_relaxed_instantiation.rs

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
