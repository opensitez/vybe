' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_reflection_get_properties_and_values
' origin: languages/vb/tests/vb/test_vb_reflection_attributes_code_gen.rs

Imports System
Imports System.Reflection

Class Configuration
    Public Property Host As String = "localhost"
    Public Property Port As Integer = 8080
End Class

Module Program
    Sub Main()
        Dim cfg As New Configuration()
        Dim props = cfg.GetType().GetProperties()
        For Each p In props
            Console.WriteLine(p.Name & "=" & p.GetValue(cfg).ToString())
        Next
    End Sub
End Module
