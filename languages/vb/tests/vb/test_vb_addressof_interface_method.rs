use super::helpers::run_vb;

#[test]
fn addressof_interface_method() {
    let out = run_vb(
        r#"
Interface IWorker
    Sub Work()
End Interface

Class Worker
    Implements IWorker
    
    Public Sub Work() Implements IWorker.Work
        Console.WriteLine("InterfaceWork")
    End Sub
End Class

Module M
    Sub Main()
        Dim w As IWorker = New Worker()
        
        ' AddressOf through an interface
        Dim act As Action = AddressOf w.Work
        act()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["InterfaceWork"]);
}
