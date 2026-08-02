' vybe-test: vb/vb_oop_inheritance/inherit_shadowed_method_access_via_base_ref
' origin: languages/vb/tests/vb/test_vb_oop_inheritance.rs

Class B
Public Sub M1()
Console.WriteLine("B")
End Sub
End Class
Class C
Inherits B
Public Shadows Sub M1()
Console.WriteLine("C")
End Sub
End Class
Module M
Sub Main()
Dim b1 As B = New C()
b1.M1()
End Sub
End Module
