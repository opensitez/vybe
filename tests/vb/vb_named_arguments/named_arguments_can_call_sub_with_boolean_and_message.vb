' vybe-test: vb/vb_named_arguments/named_arguments_can_call_sub_with_boolean_and_message
' origin: languages/vb/tests/vb/test_vb_named_arguments.rs

Module M
    Sub PrintLine(message As String, uppercase As Boolean)
        If uppercase Then
            Console.WriteLine(message & "!")
        Else
            Console.WriteLine(message)
        End If
    End Sub

    Sub Main()
        PrintLine(uppercase:=True, message:="flagged")
    End Sub
End Module
