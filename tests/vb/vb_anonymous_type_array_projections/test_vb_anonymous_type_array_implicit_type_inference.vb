' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_array_implicit_type_inference
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

Module Program
    Sub Main()
        Dim people = {
            New With {.Name = "Alice", .Score = 90},
            New With {.Name = "Bob", .Score = 85}
        }
        For Each p In people
            Console.WriteLine(p.Name & ":" & p.Score)
        Next
    End Sub
End Module
