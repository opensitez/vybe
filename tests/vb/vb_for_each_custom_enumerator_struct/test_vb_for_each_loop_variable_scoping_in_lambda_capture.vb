' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_loop_variable_scoping_in_lambda_capture
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim actions As New List(Of Action)()
        For Each item In New String() {"A", "B", "C"}
            actions.Add(Sub() Console.WriteLine(item))
        Next

        For Each act In actions
            act()
        Next
    End Sub
End Module
