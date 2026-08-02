' vybe-test: vb/vb_writeonly_properties_adv/writeonly_properties_adv
' origin: languages/vb/tests/vb/test_vb_writeonly_properties_adv.rs

Class Logger
    Private _lastMsg As String
    
    ' WriteOnly property
    Public WriteOnly Property Message As String
        Set(value As String)
            _lastMsg = value
            Console.WriteLine("Logged: " & value)
        End Set
    End Property
    
    Public Function GetLast() As String
        Return _lastMsg
    End Function
End Class

Module M
    Sub Main()
        Dim l As New Logger()
        l.Message = "Start"
        l.Message = "End"
        
        Console.WriteLine(l.GetLast())
    End Sub
End Module
