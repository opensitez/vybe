' vybe-test: vb/vb_index_out_of_range_exception/test_vb_read_only_collection_bounds_check
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"Item1"}
        Dim ro As ReadOnlyCollection(Of String) = list.AsReadOnly()
        Try
            Dim val = ro(2)
            Console.WriteLine(val)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ReadOnlyCollection ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
