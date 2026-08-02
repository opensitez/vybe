' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_foreach_can_mutate_running_total
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim items() As Integer = {1,2,3} : Dim total As Integer = 0 : For Each item In items : total += item : Next : Console.WriteLine(total) : End Sub : End Module
