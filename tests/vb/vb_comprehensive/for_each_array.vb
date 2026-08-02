' vybe-test: vb/vb_comprehensive/for_each_array
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim fruits() As String = {"apple", "banana", "cherry"}
        For Each fruit As String In fruits
            Console.WriteLine(fruit)
        Next
    End Sub
End Module
