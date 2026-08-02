' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_literal_of_booleans_can_be_counted
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim flags() As Boolean = {True, False, True} : Dim count As Integer = 0 : For Each flag In flags : If flag Then count += 1 : Next : Console.WriteLine(count) : End Sub : End Module
