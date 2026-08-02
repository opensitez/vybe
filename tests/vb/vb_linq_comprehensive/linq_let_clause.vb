' vybe-test: vb/vb_linq_comprehensive/linq_let_clause
' origin: languages/vb/tests/vb/test_vb_linq_comprehensive.rs

Module M
    Sub Main()
        Dim nums() = {1, 2, 3}
        Dim q = From n In nums
                Let sq = n * n
                Where sq > 4
                Select sq
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
