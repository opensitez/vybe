' vybe-test: vb/vb_explicit_implements/explicit_implements_different_name
' origin: languages/vb/tests/vb/test_vb_explicit_implements.rs

Interface IWorker
    Sub DoWork()
End Interface

Class Worker
    Implements IWorker
    
    ' In VB.NET, the method name doesn't have to match the interface method name,
    ' the Implements clause explicitly links them.
    Public Sub PerformTask() Implements IWorker.DoWork
        Console.WriteLine("Working")
    End Sub
End Class

Module M
    Sub Main()
        Dim w As IWorker = New Worker()
        w.DoWork()
        
        Dim c As Worker = New Worker()
        c.PerformTask()
    End Sub
End Module
