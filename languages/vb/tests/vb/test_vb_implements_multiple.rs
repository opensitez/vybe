use super::helpers::run_vb;

#[test]
fn implements_multiple() {
    let out = run_vb(
        r#"
Interface IOne
    Sub Execute()
End Interface

Interface ITwo
    Sub Execute()
End Interface

Class DualWorker
    Implements IOne, ITwo
    
    ' VB.NET allows one method to implement multiple interface methods
    Public Sub RunAll() Implements IOne.Execute, ITwo.Execute
        Console.WriteLine("Ran")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New DualWorker()
        
        Dim o1 As IOne = d
        o1.Execute()
        
        Dim o2 As ITwo = d
        o2.Execute()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Ran", "Ran"]);
}
