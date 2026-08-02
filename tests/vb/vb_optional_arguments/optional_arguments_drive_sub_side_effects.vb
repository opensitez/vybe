' vybe-test: vb/vb_optional_arguments/optional_arguments_drive_sub_side_effects
' origin: languages/vb/tests/vb/test_vb_optional_arguments.rs

Module M
    Sub AppendLine(label As String, Optional suffix As String = ".")
        Console.WriteLine(label & suffix)
    End Sub

    Sub Main()
        AppendLine("first")
        AppendLine("second", "!")
    End Sub
End Module
