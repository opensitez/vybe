' vybe-test: vb/vb_statement_separator_if/statement_separator_if
' origin: languages/vb/tests/vb/test_vb_statement_separator_if.rs

Module M
    Sub Main()
        Dim x = 10
        ' Single-line If with multiple statements separated by colons
        If x = 10 Then Console.Write("A") : Console.WriteLine("B") Else Console.WriteLine("C")
    End Sub
End Module
