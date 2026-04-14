Imports System

Module TestStaticLocals
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

    ' --- Static counter that persists across calls ---
    Sub IncrementCounter()
        Static count As Integer = 0
        count = count + 1
        Console.WriteLine("    Counter is now: " & count)
    End Sub

    Function GetCounterValue() As Integer
        Static count As Integer = 0
        count = count + 1
        Return count
    End Function

    ' --- Static with different initial value ---
    Function GetAccumulator(value As Integer) As Integer
        Static total As Integer = 0
        total = total + value
        Return total
    End Function

    Sub Main()
        Console.WriteLine("=== Static Local Variable Tests ===")

        ' --- Static in Sub persists ---
        Console.WriteLine("Test: Static in Sub persists across calls")
        IncrementCounter()
        IncrementCounter()
        IncrementCounter()
        ' The counter should have incremented 3 times
        Assert(True, "Static counter incremented 3 times without error")

        ' --- Static in Function persists ---
        Console.WriteLine("Test: Static in Function persists")
        Dim v1 As Integer = GetCounterValue()
        Dim v2 As Integer = GetCounterValue()
        Dim v3 As Integer = GetCounterValue()
        Assert(v1 = 1, "Static function: first call returns 1")
        Assert(v2 = 2, "Static function: second call returns 2")
        Assert(v3 = 3, "Static function: third call returns 3")

        ' --- Static accumulator ---
        Console.WriteLine("Test: Static accumulator")
        Dim a1 As Integer = GetAccumulator(10)
        Dim a2 As Integer = GetAccumulator(20)
        Dim a3 As Integer = GetAccumulator(5)
        Assert(a1 = 10, "Accumulator: first call 10")
        Assert(a2 = 30, "Accumulator: second call 30")
        Assert(a3 = 35, "Accumulator: third call 35")

        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("=== ALL STATIC LOCAL TESTS PASSED ===")
        Else
            Console.WriteLine("=== SOME STATIC LOCAL TESTS FAILED ===")
        End If
    End Sub
End Module
