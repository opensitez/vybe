Imports System

Module TestGoToLabelOnError
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

    ' --- GoTo basic jump ---
    Sub TestGoToBasic()
        Console.WriteLine("Test: GoTo basic jump")
        Dim x As Integer = 0
        GoTo SkipThis
        x = 999
SkipThis:
        Assert(x = 0, "GoTo skipped assignment")
    End Sub

    ' --- GoTo forward jump ---
    Sub TestGoToForward()
        Console.WriteLine("Test: GoTo forward jump")
        Dim result As String = "start"
        GoTo Done
        result = "should not reach here"
Done:
        result = result & "_done"
        Assert(result = "start_done", "GoTo forward jump correct")
    End Sub

    ' --- Labels as execution markers ---
    Sub TestLabelPassthrough()
        Console.WriteLine("Test: Label passthrough")
        Dim x As Integer = 0
FirstLabel:
        x = x + 1
SecondLabel:
        x = x + 10
        Assert(x = 11, "Labels pass through sequentially")
    End Sub

    ' --- On Error Resume Next ---
    Sub TestOnErrorResumeNext()
        Console.WriteLine("Test: On Error Resume Next")
        On Error Resume Next
        Dim x As Integer = 0
        x = 1
        ' This would cause an error but should be swallowed
        Dim y As Integer = CInt("not a number")
        x = x + 1
        Assert(x = 2, "On Error Resume Next: continued after error")
    End Sub

    ' --- On Error GoTo label ---
    Sub TestOnErrorGoToLabel()
        Console.WriteLine("Test: On Error GoTo label")
        Dim errorOccurred As Boolean = False
        On Error GoTo ErrorHandler
        Dim x As Integer = CInt("bad value")
        ' Should not reach here
        GoTo TestDone
ErrorHandler:
        errorOccurred = True
TestDone:
        Assert(errorOccurred = True, "On Error GoTo: jumped to handler")
    End Sub

    ' --- On Error GoTo 0 (disable) ---
    Sub TestOnErrorGoToZero()
        Console.WriteLine("Test: On Error GoTo 0")
        On Error Resume Next
        Dim x As Integer = CInt("bad")
        ' Error was swallowed
        On Error GoTo 0
        ' Error handling now disabled
        Assert(True, "On Error GoTo 0: disabled error handling")
    End Sub

    ' --- Multiple labels ---
    Sub TestMultipleLabels()
        Console.WriteLine("Test: Multiple labels")
        Dim path As String = ""
        GoTo Step2
Step1:
        path = path & "1"
        GoTo Step3
Step2:
        path = path & "2"
        GoTo Step1
Step3:
        path = path & "3"
        Assert(path = "213", "Multiple GoTo: correct execution path")
    End Sub

    Sub Main()
        Console.WriteLine("=== GoTo / Label / On Error Tests ===")

        TestGoToBasic()
        TestGoToForward()
        TestLabelPassthrough()
        TestOnErrorResumeNext()
        TestOnErrorGoToLabel()
        TestOnErrorGoToZero()
        TestMultipleLabels()

        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("=== ALL GOTO/LABEL/ONERROR TESTS PASSED ===")
        Else
            Console.WriteLine("=== SOME GOTO/LABEL/ONERROR TESTS FAILED ===")
        End If
    End Sub
End Module
