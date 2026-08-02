' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_deep_type_argument_substitution
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

Imports System.Collections.Generic

Interface IDataPipeline(Of TIn, TOut)
    Function Process(input As IEnumerable(Of TIn)) As List(Of TOut)
End Interface

Class StringLengthPipeline
    Implements IDataPipeline(Of String, Integer)
    Public Function Process(input As IEnumerable(Of String)) As List(Of Integer) Implements IDataPipeline(Of String, Integer).Process
        Dim res As New List(Of Integer)()
        For Each s In input
            res.Add(s.Length)
        Next
        Return res
    End Function
End Class

Module Program
    Sub Main()
        Dim p As IDataPipeline(Of String, Integer) = New StringLengthPipeline()
        Dim lengths = p.Process({"A", "BB", "CCC"})
        Console.WriteLine(String.Join(",", lengths))
    End Sub
End Module
