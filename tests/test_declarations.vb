Imports System

Module TestDeclarations
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, name As String)
        If condition Then
            passed = passed + 1
            Console.WriteLine("  PASS: " & name)
        Else
            failed = failed + 1
            Console.WriteLine("  FAIL: " & name)
        End If
    End Sub

    ' --- Interface declaration ---
    Interface IShape
        Function GetArea() As Double
        Sub Describe()
    End Interface

    ' --- Structure declaration ---
    Structure Point
        Public X As Integer
        Public Y As Integer
    End Structure

    ' --- Delegate declaration ---
    Delegate Function MathOp(a As Integer, b As Integer) As Integer

    ' --- Event declaration ---
    Event DataChanged(sender As Object, value As Integer)

    ' --- Class implementing interface (structure is treated like a class) ---
    Class Circle
        Public Radius As Double

        Sub New(r As Double)
            Radius = r
        End Sub

        Function GetArea() As Double
            Return 3.14159 * Radius * Radius
        End Function

        Sub Describe()
            Console.WriteLine("    Circle with radius " & Radius)
        End Sub
    End Class

    Sub Main()
        Console.WriteLine("=== Interface/Structure/Delegate/Event Declaration Tests ===")

        ' --- Interface parsed successfully ---
        Console.WriteLine("Test: Interface declaration parsed")
        Assert(True, "Interface IShape declared without error")

        ' --- Structure usage ---
        Console.WriteLine("Test: Structure declaration parsed")
        Dim p As New Point()
        p.X = 10
        p.Y = 20
        Assert(p.X = 10, "Structure field X = 10")
        Assert(p.Y = 20, "Structure field Y = 20")

        ' --- Delegate parsed ---
        Console.WriteLine("Test: Delegate declaration parsed")
        Assert(True, "Delegate MathOp declared without error")

        ' --- Event parsed ---
        Console.WriteLine("Test: Event declaration parsed")
        Assert(True, "Event DataChanged declared without error")

        ' --- Class with interface-like methods ---
        Console.WriteLine("Test: Class using interface pattern")
        Dim c As New Circle(5.0)
        c.Describe()
        Dim area As Double = c.GetArea()
        Assert(area > 78, "Circle area > 78")
        Assert(area < 79, "Circle area < 79")

        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("=== ALL DECLARATION TESTS PASSED ===")
        Else
            Console.WriteLine("=== SOME DECLARATION TESTS FAILED ===")
        End If
    End Sub
End Module
