' vybe-test: vb/vb_spec_arrays_collections/array_spec_array_bounds_can_drive_descending_loop
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim items() As Integer = {1,2,3} : Dim total As Integer = 0 : For i As Integer = UBound(items) To LBound(items) Step -1 : total += items(i) : Next : Console.WriteLine(total) : End Sub : End Module
