' vybe-test: vb/vb_linq_comprehensive/linq_where_select
' origin: languages/vb/tests/vb/test_vb_linq_comprehensive.rs

Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 5, 6}
        Dim q = From n In nums Where n Mod 2 = 0 Select n * 2
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
