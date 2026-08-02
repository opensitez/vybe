' vybe-test: vb/generators/iterator_function_body_stays_lazy
' origin: languages/vb/tests/vb/test_generators.rs

Module Program
    Function Loud()
        Console.WriteLine("bad")
        Yield 1
    End Function

    Sub Main()
        Dim g = Loud()
        Console.WriteLine("ok")
    End Sub
End Module
