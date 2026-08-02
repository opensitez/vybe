' vybe-test: vb/vb_paramarray_byval/paramarray_byval
' origin: languages/vb/tests/vb/test_vb_paramarray_byval.rs

Module M
    ' ParamArray is always implicitly ByVal in modern VB, but you can explicitly specify it
    Sub PrintAll(ByVal ParamArray items() As Integer)
        Console.WriteLine(items.Length)
        For Each item In items
            Console.WriteLine(item)
        Next
    End Sub

    Sub Main()
        PrintAll(10, 20, 30)
    End Sub
End Module
