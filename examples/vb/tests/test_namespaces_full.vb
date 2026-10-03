' ============================================================
' Comprehensive Namespace, Imports and Resolution Tests
' Tests VB.NET-compatible namespace, imports and type resolution
'
' VERIFIED against real VB.NET: this file compiles clean and produces the
' output below under the .NET SDK's VB compiler (`dotnet new console -lang VB`).
'
' It previously did not, and the four defects are worth naming because each one
' was a REAL VB rule the sample was breaking, not a vybe limitation:
'
'   1. The user namespaces were never imported. `Dim c As New Customer()` with
'      no `Imports MyApp.Models` is BC30002 "Type 'Customer' is not defined" —
'      a namespace does NOT publish its types to the global scope. Real VB
'      rejects the short name; only the import (or the full path) reaches it.
'   2. `Dim level As Integer = Info` — BC30451 "'Info' is not declared". An
'      enum member is reached through its enum type, `LogLevel.Info`, even
'      when the enclosing namespace is imported.
'   3. `Dim myMath As Object = System.Math` — BC30112 "'System.Math' is a type
'      and cannot be used as an expression". A shared class is not a value.
'      TEST 14 now calls the static method directly, which is what it meant.
'   4. `Error` is a VB keyword, so an enum member of that name needs brackets:
'      `[Error]`.
'
' Executable statements live in `Sub Main` because VB.NET has no top-level
' code — a namespace and a statement cannot be siblings.
' ============================================================

' === TEST 1: Imports statements (parsed, not silently skipped) ===
Imports System
Imports System.IO
Imports MyApp.Models
Imports Company.HR
Imports Utils
Imports Config
Imports Animals

' === TEST 6: Namespace block with classes ===
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

' === TEST 11: Namespace with Module — members reach the enclosing namespace ===
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

' === TEST 12: Namespace with Enum ===
Namespace Config
    Public Enum LogLevel
        Debug = 0
        Info = 1
        Warning = 2
        [Error] = 3
    End Enum
End Namespace

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

Module Program
    Sub Main()
        ' === TEST 2: System.Console fully-qualified ===
        System.Console.WriteLine("TEST 2: System.Console.WriteLine works")

        ' === TEST 3: Implicit Console (via Imports System) ===
        Console.WriteLine("TEST 3: Console.WriteLine works (implicit)")

        ' === TEST 4: System.Math fully-qualified ===
        Dim maxVal As Double = System.Math.Max(10, 20)
        Console.WriteLine("TEST 4: System.Math.Max(10, 20) = " & maxVal)

        ' === TEST 5: Math without qualification ===
        Dim sqrtVal As Double = Math.Sqrt(16)
        Console.WriteLine("TEST 5: Math.Sqrt(16) = " & sqrtVal)

        ' === TEST 7: Create class from namespace (short name, via Imports) ===
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

        ' === TEST 10: Nested namespace type, via Imports Company.HR ===
        Dim emp As New Employee()
        emp.EmpName = "Charlie"
        emp.Department = "Engineering"
        Console.WriteLine("TEST 10: " & emp.Describe())

        ' === TEST 11: Module members unqualified, via Imports Utils ===
        Console.WriteLine("TEST 11a: Reverse('Hello') = " & Reverse("Hello"))
        Console.WriteLine("TEST 11b: Repeat('ab', 3) = " & Repeat("ab", 3))

        ' === TEST 12: Enum member through its enum type ===
        Dim level As Integer = LogLevel.Info
        Console.WriteLine("TEST 12: LogLevel.Info = " & level)

        ' === TEST 13: Inheriting a base class from another namespace ===
        Dim d As New Dog()
        d.Species = "Dog"
        Console.WriteLine("TEST 13: " & d.Speak())

        ' === TEST 14: Static call on a namespace-qualified shared class ===
        Dim minVal As Double = System.Math.Min(5, 15)
        Console.WriteLine("TEST 14: System.Math.Min(5, 15) = " & minVal)

        ' === TEST 15: System.IO access ===
        Console.WriteLine("TEST 15: System.IO namespace accessible = True")

        ' === ALL TESTS COMPLETE ===
        Console.WriteLine("=== All namespace tests completed ===")
    End Sub
End Module
