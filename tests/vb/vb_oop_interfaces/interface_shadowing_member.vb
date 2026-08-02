' vybe-test: vb/vb_oop_interfaces/interface_shadowing_member
' origin: languages/vb/tests/vb/test_vb_oop_interfaces.rs

Interface I1
Sub M()
End Interface
Interface I2
Inherits I1
Shadows Sub M()
End Interface
Class C
Implements I2
Public Sub M1() Implements I1.M
Console.WriteLine("1")
End Sub
Public Sub M2() Implements I2.M
Console.WriteLine("2")
End Sub
End Class
Module M
Sub Main()
Dim c1 As I1 = New C()
c1.M()
End Sub
End Module
