' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_loop_variable_explicit_type_widening
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Module Program
    Sub Main()
        Dim bytes As Byte() = {1, 2, 3}
        ' Explicitly typed loop variable Double widens from Byte!
        For Each val As Double In bytes
            Console.WriteLine(val.ToString("F1"))
        Next
    End Sub
End Module
