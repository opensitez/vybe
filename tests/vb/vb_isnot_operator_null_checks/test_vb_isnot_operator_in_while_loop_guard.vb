' vybe-test: vb/vb_isnot_operator_null_checks/test_vb_isnot_operator_in_while_loop_guard
' origin: languages/vb/tests/vb/test_vb_isnot_operator_null_checks.rs

Module Program
    Class Node
        Public Value As Integer
        Public NextNode As Node
    End Class

    Sub Main()
        Dim head As New Node With {.Value = 1, .NextNode = New Node With {.Value = 2}}
        Dim current As Node = head
        Dim sum = 0
        While current IsNot Nothing
            sum += current.Value
            current = current.NextNode
        End While
        Console.WriteLine(sum)
    End Sub
End Module
