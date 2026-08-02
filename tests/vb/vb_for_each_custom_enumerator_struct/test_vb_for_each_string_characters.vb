' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_string_characters
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Module Program
    Sub Main()
        Dim text = "Vybe"
        For Each ch As Char In text
            Console.WriteLine(ch)
        Next
    End Sub
End Module
