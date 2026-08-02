' vybe-test: vb/vb_linq_comprehensive/linq_group_by_comprehensive
' origin: languages/vb/tests/vb/test_vb_linq_comprehensive.rs

Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 5}
        Dim q = From n In nums Group By IsEven = (n Mod 2 = 0) Into Group Select IsEven, Group
        For Each g In q
            Console.WriteLine(g.IsEven)
            Console.WriteLine(g.Group.Count())
        Next
    End Sub
End Module
