' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_empty_collection_does_not_execute_body
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer)()
        Dim executed = False
        For Each item In list
            executed = True
        Next
        Console.WriteLine(executed)
    End Sub
End Module
