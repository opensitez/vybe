' vybe-test: vb/vb_builtins/select_case_string
' origin: languages/vb/tests/vb/vb_builtins_test.rs

Module Program
    Sub Main()
        Dim color As String = "red"
        Select Case color
            Case "blue"
                Console.WriteLine("sky")
            Case "red"
                Console.WriteLine("fire")
            Case "green"
                Console.WriteLine("grass")
        End Select
    End Sub
End Module
