' vybe-test: vb/vb_oop_interfaces/interface_inheritance
' origin: languages/vb/tests/vb/test_vb_oop_interfaces.rs

Interface I1
Sub M1()
End Interface
Interface I2
Inherits I1
Sub M2()
End Interface
Class C
Implements I2
Public Sub M1() Implements I2.M1
Console.WriteLine("1")
End Sub
Public Sub M2() Implements I2.M2
Console.WriteLine("2")
End Sub
End Class
Module M
Sub Main()
Dim c1 As New C()
c1.M1()
End Sub
End Module
