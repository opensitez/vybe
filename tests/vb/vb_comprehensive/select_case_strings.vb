' vybe-test: vb/vb_comprehensive/select_case_strings
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim color As String = "green"
        Select Case color
            Case "red"
                Console.WriteLine("stop")
            Case "green"
                Console.WriteLine("go")
            Case "yellow"
                Console.WriteLine("caution")
        End Select
    End Sub
End Module
