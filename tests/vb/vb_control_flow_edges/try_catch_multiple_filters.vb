' vybe-test: vb/vb_control_flow_edges/try_catch_multiple_filters
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Imports System

Module M
    Sub Main()
        Dim code = 2
        Try
            Throw New Exception("Err")
        Catch ex As Exception When code = 1
            Console.WriteLine("One")
        Catch ex As Exception When code = 2
            Console.WriteLine("Two")
        End Try
    End Sub
End Module
