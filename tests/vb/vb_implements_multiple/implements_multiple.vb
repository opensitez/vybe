' vybe-test: vb/vb_implements_multiple/implements_multiple
' origin: languages/vb/tests/vb/test_vb_implements_multiple.rs

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
