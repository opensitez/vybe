' vybe-test: vb/vb_oop_attributes_events/event_removehandler
' origin: languages/vb/tests/vb/test_vb_oop_attributes_events.rs

Class C: Public Event E(): Public Sub DoE(): RaiseEvent E(): End Sub: End Class: Module M: Sub Main(): Dim obj As New C(): Dim h As System.Action = Sub() Console.WriteLine("X"): AddHandler obj.E, h: RemoveHandler obj.E, h: obj.DoE(): Console.WriteLine("Done"): End Sub: End Module
