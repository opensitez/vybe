Module MainModule
    Sub Main()
        ' Test Inheritance
        Dim d As New Derived
        d.BaseField = 10
        d.DerivedField = 20
        Console.WriteLine("BaseField: " & d.BaseField) ' Should be 10
        Console.WriteLine("DerivedField: " & d.DerivedField) ' Should be 20
        d.BaseMethod() ' Should print "Base Method"
        d.DerivedMethod() ' Should print "Derived Method"

        ' Test Partial Classes
        Dim p As New PartialClass
        p.Part1Method()
        p.Part2Method()
    End Sub
End Module

Class Base
    Public BaseField As Integer
    Public Sub BaseMethod()
        Console.WriteLine("Base Method")
    End Sub
End Class

Class Derived
    Inherits Base
    Public DerivedField As Integer
    Public Sub DerivedMethod()
        Console.WriteLine("Derived Method")
    End Sub
End Class

Partial Public Class PartialClass
    Public Sub Part1Method()
        Console.WriteLine("Part 1")
    End Sub
End Class

Partial Public Class PartialClass
    Public Sub Part2Method()
        Console.WriteLine("Part 2")
    End Sub
End Class
