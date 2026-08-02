' vybe-test: vb/vb_spec_arrays_collections/array_spec_dictionary_can_iterate_keys_with_foreach
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

Module M : Sub Main() : Dim d As New Dictionary(Of String, Integer) : d.Add("a",1) : d.Add("b",2) : Dim total As Integer = 0 : For Each key In d.Keys : total += d.Item(key) : Next : Console.WriteLine(total) : End Sub : End Module
