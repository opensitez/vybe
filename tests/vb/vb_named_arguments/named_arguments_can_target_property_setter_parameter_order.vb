' vybe-test: vb/vb_named_arguments/named_arguments_can_target_property_setter_parameter_order
' origin: languages/vb/tests/vb/test_vb_named_arguments.rs

Class Counter
    Public Value As Integer

    Public Sub Apply(amount As Integer, repeatCount As Integer)
        For i As Integer = 1 To repeatCount
            Value = Value + amount
        Next
    End Sub
End Class

Module M
    Sub Main()
        Dim counter As New Counter()
        counter.Apply(repeatCount:=3, amount:=4)
        Console.WriteLine(counter.Value)
    End Sub
End Module
