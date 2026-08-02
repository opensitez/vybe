' vybe-test: vb/vb_static_locals/static_local_in_sub_tracks_invocation_count
' origin: languages/vb/tests/vb/test_vb_static_locals.rs

Module M
    Sub Report()
        Static callCount As Integer = 0
        callCount = callCount + 1
        Console.WriteLine(callCount)
    End Sub

    Sub Main()
        Report()
        Report()
        Report()
    End Sub
End Module
