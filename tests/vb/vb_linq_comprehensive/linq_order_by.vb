' vybe-test: vb/vb_linq_comprehensive/linq_order_by
' origin: languages/vb/tests/vb/test_vb_linq_comprehensive.rs

Module M
    Sub Main()
        Dim names() = {"Charlie", "Alice", "Bob"}
        Dim q = From n In names Order By n Descending Select n
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
