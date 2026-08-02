' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_method_in_class_instance
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Class DataPipeline
    Private data As String() = {"X", "Y", "Z"}

    Public Iterator Function GetFilteredData() As IEnumerable(Of String)
        For Each d In data
            If d <> "Y" Then Yield d
        Next
    End Function
End Class

Module Program
    Sub Main()
        Dim pipe As New DataPipeline()
        Console.WriteLine(String.Join("", pipe.GetFilteredData()))
    End Sub
End Module
