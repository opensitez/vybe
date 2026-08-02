' vybe-test: vb/vb_shared_vs_static/shared_vs_static
' origin: languages/vb/tests/vb/test_vb_shared_vs_static.rs

Class Counter
    ' Shared is used for class-level static variables
    Public Shared Value As Integer = 0
    
    Public Sub Increment()
        Value += 1
    End Sub
End Class

Module M
    Sub TestStatic()
        ' Static is used for local variables
        Static localVal As Integer = 0
        localVal += 1
        Console.WriteLine(localVal)
    End Sub

    Sub Main()
        Dim c1 As New Counter()
        Dim c2 As New Counter()
        
        c1.Increment()
        c2.Increment()
        
        Console.WriteLine(Counter.Value)
        
        TestStatic()
        TestStatic()
    End Sub
End Module
