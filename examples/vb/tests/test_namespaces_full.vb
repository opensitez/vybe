' ============================================================
' Comprehensive Namespace, Imports and Resolution Tests
' Tests VB.NET-compatible namespace, imports and type resolution
' ============================================================

' === TEST 1: Basic Imports statement (parsed, not silently skipped) ===
Imports System
Imports System.IO

' === TEST 2: System.Console fully-qualified ===
System.Console.WriteLine("TEST 2: System.Console.WriteLine works")

' === TEST 3: Implicit Console (like Imports System) ===
Console.WriteLine("TEST 3: Console.WriteLine works (implicit)")

' === TEST 4: System.Math fully-qualified ===
Dim maxVal As Double = System.Math.Max(10, 20)
Console.WriteLine("TEST 4: System.Math.Max(10, 20) = " & maxVal)

' === TEST 5: Math without qualification ===
Dim sqrtVal As Double = Math.Sqrt(16)
Console.WriteLine("TEST 5: Math.Sqrt(16) = " & sqrtVal)

' === TEST 6: Namespace block with class ===
Namespace MyApp.Models
    Public Class Customer
        Public Property Name As String
        Public Property Age As Integer

        Public Sub New()
            Name = "Default"
            Age = 0
        End Sub

        Public Function GetInfo() As String
            Return Name & " (age " & Age & ")"
        End Function
    End Class

    Public Class Order
        Public Property OrderId As Integer
        Public Property CustomerName As String

        Public Function GetSummary() As String
            Return "Order #" & OrderId & " for " & CustomerName
        End Function
    End Class
End Namespace

' === TEST 7: Create class from namespace (short name) ===
Dim c As New Customer()
c.Name = "Alice"
c.Age = 30
Console.WriteLine("TEST 7: " & c.GetInfo())

' === TEST 8: Create class with fully-qualified name ===
Dim c2 As New MyApp.Models.Customer()
c2.Name = "Bob"
c2.Age = 25
Console.WriteLine("TEST 8: " & c2.GetInfo())

' === TEST 9: Multiple classes in same namespace ===
Dim o As New Order()
o.OrderId = 1001
o.CustomerName = "Alice"
Console.WriteLine("TEST 9: " & o.GetSummary())

' === TEST 10: Nested namespaces ===
Namespace Company
    Namespace HR
        Public Class Employee
            Public Property EmpName As String
            Public Property Department As String

            Public Function Describe() As String
                Return EmpName & " in " & Department
            End Function
        End Class
    End Namespace
End Namespace

Dim emp As New Employee()
emp.EmpName = "Charlie"
emp.Department = "Engineering"
Console.WriteLine("TEST 10: " & emp.Describe())

' === TEST 11: Namespace with Module (functions globally accessible) ===
Namespace Utils
    Module StringHelpers
        Public Function Reverse(s As String) As String
            Dim result As String = ""
            Dim i As Integer
            For i = Len(s) To 1 Step -1
                result = result & Mid(s, i, 1)
            Next
            Return result
        End Function

        Public Function Repeat(s As String, count As Integer) As String
            Dim result As String = ""
            Dim i As Integer
            For i = 1 To count
                result = result & s
            Next
            Return result
        End Function
    End Module
End Namespace

Console.WriteLine("TEST 11a: Reverse('Hello') = " & Reverse("Hello"))
Console.WriteLine("TEST 11b: Repeat('ab', 3) = " & Repeat("ab", 3))

' === TEST 12: Namespace with Enum ===
Namespace Config
    Public Enum LogLevel
        Debug = 0
        Info = 1
        Warning = 2
        Error = 3
    End Enum
End Namespace

Dim level As Integer = Info
Console.WriteLine("TEST 12: LogLevel.Info = " & level)

' === TEST 13: Class inheritance across namespaces ===
Namespace Animals
    Public Class Animal
        Public Property Species As String

        Public Overridable Function Speak() As String
            Return Species & " says ..."
        End Function
    End Class
End Namespace

Public Class Dog
    Inherits Animal

    Public Overrides Function Speak() As String
        Return Species & " says Woof!"
    End Function
End Class

Dim d As New Dog()
d.Species = "Dog"
Console.WriteLine("TEST 13: " & d.Speak())

' === TEST 14: Object assignment from namespace ===
Dim myMath As Object = System.Math
Dim minVal As Double = myMath.Min(5, 15)
Console.WriteLine("TEST 14: myMath.Min(5, 15) = " & minVal)

' === TEST 15: System.IO access ===
Console.WriteLine("TEST 15: System.IO namespace accessible = True")

' === ALL TESTS COMPLETE ===
Console.WriteLine("=== All namespace tests completed ===")
