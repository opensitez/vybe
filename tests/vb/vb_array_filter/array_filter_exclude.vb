' vybe-test: vb/vb_array_filter/array_filter_exclude
' origin: languages/vb/tests/vb/test_vb_array_filter.rs

Module M
    Sub Main()
        Dim source As String() = {"apple", "banana", "apricot", "cherry"}
        
        ' Filter with Include=False returns elements that DO NOT contain the match
        Dim result As String() = Filter(source, "ap", False)
        
        For Each item In result
            Console.WriteLine(item)
        Next
    End Sub
End Module
