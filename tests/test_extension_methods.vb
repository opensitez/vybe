Imports System
Imports System.Runtime.CompilerServices

' Extension method module
Module StringExtensions
    <Extension()>
    Public Function Reverse(s As String) As String
        Return "REVERSED:" & s
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(s As String) As Boolean
        Return s Is Nothing OrElse s.Length = 0
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Sub PrintUpper(s As String)
        Console.WriteLine(s.ToUpper())
    End Sub
End Module

Module Program
    Sub Main()
        ' Test 1: Extension function call
        Dim word As String = "Hello"
        Dim reversed As String = word.Reverse()
        Console.WriteLine("TEST 1: " & reversed)

        ' Test 2: Extension function with fully qualified attribute
        Dim empty As String = ""
        Console.WriteLine("TEST 2: " & empty.IsNullOrEmpty())

        ' Test 3: Extension sub call
        Dim msg As String = "hello world"
        Console.Write("TEST 3: ")
        msg.PrintUpper()

        ' Test 4: Extension on literal-like expression
        Dim greeting As String = "abcde"
        Console.WriteLine("TEST 4: " & greeting.Reverse())
    End Sub
End Module
