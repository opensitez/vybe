' vybe-test: vb/vb_typeof_isnot/typeof_isnot
' origin: languages/vb/tests/vb/test_vb_typeof_isnot.rs

Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module M
    Sub Main()
        Dim obj As Object = New Animal()
        
        ' VB.NET allows TypeOf ... IsNot
        If TypeOf obj IsNot Dog Then
            Console.WriteLine("Not a dog")
        End If
        
        Dim dogObj As Object = New Dog()
        If TypeOf dogObj IsNot Dog Then
            Console.WriteLine("This shouldn't print")
        Else
            Console.WriteLine("Is a dog")
        End If
    End Sub
End Module
