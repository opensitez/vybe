' vybe-test: vb/vb_optional_object_nothing/optional_object_nothing
' origin: languages/vb/tests/vb/test_vb_optional_object_nothing.rs

Module M
    ' Optional Object parameter defaulting to Nothing
    Sub DoWork(Optional obj As Object = Nothing)
        Console.WriteLine(obj Is Nothing)
    End Sub

    Sub Main()
        DoWork()
        DoWork(New Object())
    End Sub
End Module
