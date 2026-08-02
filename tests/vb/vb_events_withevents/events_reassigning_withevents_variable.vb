' vybe-test: vb/vb_events_withevents/events_reassigning_withevents_variable
' origin: languages/vb/tests/vb/test_vb_events_withevents.rs

Class Worker
    Public ID As Integer
    Public Event Working(id As Integer)
    
    Public Sub New(i As Integer)
        ID = i
    End Sub
    
    Public Sub Work()
        RaiseEvent Working(ID)
    End Sub
End Class

Module M
    Private WithEvents ActiveWorker As Worker
    
    Private Sub OnWorking(id As Integer) Handles ActiveWorker.Working
        Console.WriteLine("Worker " & id & " is working")
    End Sub
    
    Sub Main()
        Dim w1 As New Worker(1)
        Dim w2 As New Worker(2)
        
        ActiveWorker = w1
        w1.Work()
        
        ' Reassigning automatically unhooks w1 and hooks w2
        ActiveWorker = w2
        w1.Work() ' Should NOT trigger the handler
        w2.Work() ' SHOULD trigger the handler
    End Sub
End Module
