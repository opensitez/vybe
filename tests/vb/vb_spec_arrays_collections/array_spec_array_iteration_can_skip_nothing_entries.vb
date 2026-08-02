' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_iteration_can_skip_nothing_entries
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim items() As String = {"a", Nothing, "c"} : Dim s As String = "" : For Each item In items : If Not IsNothing(item) Then s &= item : Next : Console.WriteLine(s) : End Sub : End Module
