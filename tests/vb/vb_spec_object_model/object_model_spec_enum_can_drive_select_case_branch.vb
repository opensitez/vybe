' vybe-test: vb/vb_spec_object_model/object_model_spec_enum_can_drive_select_case_branch
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

Enum Tone
    Low
    High
End Enum
Module M
    Sub Main()
        Dim tone As Tone = Tone.High
        Select Case tone
            Case Tone.Low
                Console.WriteLine("low")
            Case Tone.High
                Console.WriteLine("high")
        End Select
    End Sub
End Module
