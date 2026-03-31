Public Class TestClass
    Private _flag As Boolean = False

    Public Sub New()
        Console.WriteLine("Constructor: _flag = " & _flag)
    End Sub

    Public Sub TestIfNot()
        Console.WriteLine("Before If: _flag = " & _flag)
        If Not _flag Then
            Console.WriteLine("IF branch: _flag was False")
            _flag = True
        Else
            Console.WriteLine("ELSE branch: _flag was True")
            _flag = False
        End If
        Console.WriteLine("After If: _flag = " & _flag)
    End Sub
End Class

Module Module1
    Sub Main()
        Dim obj As New TestClass()
        Console.WriteLine("--- First call ---")
        obj.TestIfNot()
        Console.WriteLine("--- Second call ---")
        obj.TestIfNot()
    End Sub
End Module
