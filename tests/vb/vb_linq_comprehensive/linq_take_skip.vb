' vybe-test: vb/vb_linq_comprehensive/linq_take_skip
' origin: languages/vb/tests/vb/test_vb_linq_comprehensive.rs

Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 5}
        Dim q = From n In nums Skip 2 Take 2 Select n
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
