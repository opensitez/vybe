' vybe-test: vb/vb_linq_comprehensive/linq_skip_while
' origin: languages/vb/tests/vb/test_vb_linq_comprehensive.rs

Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 1, 2}
        Dim q = From n In nums Skip While n < 4 Select n
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
