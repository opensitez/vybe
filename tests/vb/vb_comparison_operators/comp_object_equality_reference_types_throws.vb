' vybe-test: vb/vb_comparison_operators/comp_object_equality_reference_types_throws
' origin: languages/vb/tests/vb/test_vb_comparison_operators.rs

Option Strict Off
Class C
End Class
Module M
Sub Main()
Dim c1 As Object = New C()
Dim c2 As Object = New C()
Try
Console.WriteLine(c1 = c2)
Catch
Console.WriteLine("Caught")
End Try
End Sub
End Module
