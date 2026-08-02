' vybe-test: vb/vb_delegates_addressof/event_remove_handler
' origin: languages/vb/tests/vb/test_vb_delegates_addressof.rs

Class C
Public Event E()
Public Sub Raise()
RaiseEvent E()
End Sub
End Class
Module M
Sub Handler()
Console.WriteLine("Handled")
End Sub
Sub Main()
Dim c1 As New C()
AddHandler c1.E, AddressOf Handler
RemoveHandler c1.E, AddressOf Handler
c1.Raise()
Console.WriteLine("OK")
End Sub
End Module
