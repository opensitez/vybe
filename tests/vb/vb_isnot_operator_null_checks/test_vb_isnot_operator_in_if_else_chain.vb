' vybe-test: vb/vb_isnot_operator_null_checks/test_vb_isnot_operator_in_if_else_chain
' origin: languages/vb/tests/vb/test_vb_isnot_operator_null_checks.rs

Module Program
    Sub Main()
        Dim obj As Object = "Sample"
        If obj IsNot Nothing Then
            Console.WriteLine("Object Exists: " & obj.ToString())
        Else
            Console.WriteLine("Object Is Null")
        End If
    End Sub
End Module
