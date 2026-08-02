' vybe-test: vb/vb_builtins/select_case_basic
' origin: languages/vb/tests/vb/vb_builtins_test.rs

Module Program
    Sub Main()
        Dim x As Integer = 2
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case 2
                Console.WriteLine("two")
            Case 3
                Console.WriteLine("three")
        End Select
    End Sub
End Module
