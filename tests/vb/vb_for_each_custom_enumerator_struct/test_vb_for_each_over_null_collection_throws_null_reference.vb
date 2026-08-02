' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_over_null_collection_throws_null_reference
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As List(Of Integer) = Nothing
        Try
            For Each item In list
                Console.WriteLine(item)
            Next
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught on For Each Null")
        End Try
    End Sub
End Module
