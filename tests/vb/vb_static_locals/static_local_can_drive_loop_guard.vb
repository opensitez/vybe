' vybe-test: vb/vb_static_locals/static_local_can_drive_loop_guard
' origin: languages/vb/tests/vb/test_vb_static_locals.rs

Module M
    Function NextValue() As Integer
        Static total As Integer = 0
        total = total + 2
        Return total
    End Function

    Sub Main()
        Do While NextValue() < 7
            Console.WriteLine("loop")
        Loop
        Console.WriteLine(NextValue())
    End Sub
End Module
