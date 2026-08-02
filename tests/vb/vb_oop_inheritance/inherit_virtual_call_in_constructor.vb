' vybe-test: vb/vb_oop_inheritance/inherit_virtual_call_in_constructor
' origin: languages/vb/tests/vb/test_vb_oop_inheritance.rs

Class B
Public Sub New()
M1()
End Sub
Public Overridable Sub M1()
Console.WriteLine("B")
End Sub
End Class
Class C
Inherits B
Public Overrides Sub M1()
Console.WriteLine("C")
End Sub
End Class
Module M
Sub Main()
Dim c1 As New C()
End Sub
End Module
