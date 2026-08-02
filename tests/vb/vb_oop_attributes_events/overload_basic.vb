' vybe-test: vb/vb_oop_attributes_events/overload_basic
' origin: languages/vb/tests/vb/test_vb_oop_attributes_events.rs

Module M: Sub Test(x As Integer): Console.WriteLine("I"): End Sub: Sub Test(x As String): Console.WriteLine("S"): End Sub: Sub Main(): Test("A"): End Sub: End Module
