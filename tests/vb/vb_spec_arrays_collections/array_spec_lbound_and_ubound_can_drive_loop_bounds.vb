' vybe-test: vb/vb_spec_arrays_collections/array_spec_lbound_and_ubound_can_drive_loop_bounds
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim items() As Integer = {2,4,6} : Dim total As Integer = 0 : For i As Integer = LBound(items) To UBound(items) : total += items(i) : Next : Console.WriteLine(total) : End Sub : End Module
