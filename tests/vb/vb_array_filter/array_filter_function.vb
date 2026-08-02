' vybe-test: vb/vb_array_filter/array_filter_function
' origin: languages/vb/tests/vb/test_vb_array_filter.rs

Module M
    Sub Main()
        Dim source As String() = {"apple", "banana", "apricot", "cherry"}
        
        ' Filter returns a new array with elements containing the match string
        Dim result As String() = Filter(source, "ap")
        
        For Each item In result
            Console.WriteLine(item)
        Next
    End Sub
End Module
