' vybe-test: vb/vb_spec_arrays_collections/array_spec_for_index_loop_can_write_each_array_slot
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim values(2) As Integer : For i As Integer = 0 To 2 : values(i)=i+1 : Next : Console.WriteLine(values(2)) : End Sub : End Module
