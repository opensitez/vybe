' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_of_enums_can_be_compared_in_loop
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Enum Tone : Low : High : End Enum : Module M : Sub Main() : Dim tones() As Tone = {Tone.Low, Tone.High} : Dim count As Integer = 0 : For Each t In tones : If t = Tone.High Then count += 1 : Next : Console.WriteLine(count) : End Sub : End Module
