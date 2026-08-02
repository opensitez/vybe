' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_iteration_can_build_delimited_string
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim items() As String = {"a","b","c"} : Dim s As String = "" : For Each item In items : If s <> "" Then s &= "|" : s &= item : Next : Console.WriteLine(s) : End Sub : End Module
