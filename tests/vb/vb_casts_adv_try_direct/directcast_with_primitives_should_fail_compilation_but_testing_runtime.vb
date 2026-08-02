' vybe-test: vb/vb_casts_adv_try_direct/directcast_with_primitives_should_fail_compilation_but_testing_runtime
' origin: languages/vb/tests/vb/test_vb_casts_adv_try_direct.rs

Module M
    Sub Main()
        Dim obj As Object = 42
        
        ' DirectCast requires exact type match for value types boxed in Object
        Dim i As Integer = DirectCast(obj, Integer)
        Console.WriteLine(i)
        
        Try
            ' This throws InvalidCastException at runtime
            Dim d As Double = DirectCast(obj, Double)
            Console.WriteLine(d)
        Catch ex As Exception
            Console.WriteLine("Cast Failed")
        End Try
    End Sub
End Module
