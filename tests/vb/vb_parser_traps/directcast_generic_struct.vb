' vybe-test: vb/vb_parser_traps/directcast_generic_struct
' origin: languages/vb/tests/vb/test_vb_parser_traps.rs

Class Tester
    Public Sub Process(Of T As Structure)(obj As Object)
        Try
            Dim c = DirectCast(obj, T)
            Console.WriteLine("Cast Success")
        Catch
            Console.WriteLine("Cast Failed")
        End Try
    End Sub
End Class

Module M
    Sub Main()
        Dim t As New Tester()
        t.Process(Of Integer)("String")
    End Sub
End Module
