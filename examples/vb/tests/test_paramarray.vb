Imports System

Module TestParamArray
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

    ' --- ParamArray on Sub ---
    Sub PrintAll(ParamArray items() As String)
        Dim result As String = ""
        For Each item As String In items
            If result <> "" Then result = result & ", "
            result = result & item
        Next
        Console.WriteLine("    Items: " & result)
    End Sub

    ' --- ParamArray on Function ---
    Function SumAll(ParamArray numbers() As Integer) As Integer
        Dim total As Integer = 0
        For Each n As Integer In numbers
            total = total + n
        Next
        Return total
    End Function

    ' --- ParamArray with a leading fixed parameter ---
    Function FormatMessage(prefix As String, ParamArray parts() As String) As String
        Dim result As String = prefix & ": "
        For Each p As String In parts
            result = result & p & " "
        Next
        Return result
    End Function

    ' --- ParamArray returning count ---
    Function CountArgs(ParamArray args() As Object) As Integer
        Return args.Length
    End Function

    Sub Main()
        Console.WriteLine("=== ParamArray Tests ===")

        ' --- ParamArray Sub with multiple args ---
        Console.WriteLine("Test: ParamArray Sub")
        PrintAll("Hello", "World", "VB")
        Assert(True, "ParamArray Sub: called without error")

        ' --- ParamArray Function sum ---
        Console.WriteLine("Test: ParamArray Function sum")
        Dim s1 As Integer = SumAll(1, 2, 3)
        Assert(s1 = 6, "SumAll(1,2,3) = 6")
        
        Dim s2 As Integer = SumAll(10, 20, 30, 40)
        Assert(s2 = 100, "SumAll(10,20,30,40) = 100")

        ' --- ParamArray with single arg ---
        Console.WriteLine("Test: ParamArray single arg")
        Dim s3 As Integer = SumAll(42)
        Assert(s3 = 42, "SumAll(42) = 42")

        ' --- ParamArray with no variadic args ---
        Console.WriteLine("Test: ParamArray no args")
        Dim s4 As Integer = SumAll()
        Assert(s4 = 0, "SumAll() = 0")

        ' --- ParamArray with fixed + variadic ---
        Console.WriteLine("Test: ParamArray with fixed param")
        Dim msg As String = FormatMessage("Error", "file", "not", "found")
        Console.WriteLine("    Message: " & msg)
        Assert(True, "FormatMessage with ParamArray called")

        ' --- ParamArray counting ---
        Console.WriteLine("Test: ParamArray counting")
        Assert(CountArgs() = 0, "CountArgs() = 0")
        Assert(CountArgs("a") = 1, "CountArgs(a) = 1")
        Assert(CountArgs("a", "b", "c") = 3, "CountArgs(a,b,c) = 3")

        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("=== ALL PARAMARRAY TESTS PASSED ===")
        Else
            Console.WriteLine("=== SOME PARAMARRAY TESTS FAILED ===")
        End If
    End Sub
End Module
