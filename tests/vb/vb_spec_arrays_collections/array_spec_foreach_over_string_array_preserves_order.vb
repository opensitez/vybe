' vybe-test: vb/vb_spec_arrays_collections/array_spec_foreach_over_string_array_preserves_order
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim items() As String = {"a","b","c"} : Dim s As String = "" : For Each item In items : s &= item : Next : Console.WriteLine(s) : End Sub : End Module
