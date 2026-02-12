Module InterpolatedStringTests
    Sub Main()
        ' Basic interpolated string
        Dim name As String = "World"
        Dim result As String = $"Hello {name}!"
        Console.WriteLine("Test 1: " & result)
        If result = "Hello World!" Then
            Console.WriteLine("PASS: Basic interpolation")
        Else
            Console.WriteLine("FAIL: Basic interpolation, got: " & result)
        End If

        ' Multiple interpolations
        Dim firstName As String = "John"
        Dim lastName As String = "Doe"
        Dim greeting As String = $"Hello {firstName} {lastName}!"
        Console.WriteLine("Test 2: " & greeting)
        If greeting = "Hello John Doe!" Then
            Console.WriteLine("PASS: Multiple interpolations")
        Else
            Console.WriteLine("FAIL: Multiple interpolations, got: " & greeting)
        End If

        ' Expression in interpolation
        Dim x As Integer = 5
        Dim y As Integer = 3
        Dim mathResult As String = $"Sum is {x + y}"
        Console.WriteLine("Test 3: " & mathResult)
        If mathResult = "Sum is 8" Then
            Console.WriteLine("PASS: Expression interpolation")
        Else
            Console.WriteLine("FAIL: Expression interpolation, got: " & mathResult)
        End If

        ' Interpolation with method call
        Dim text As String = "hello"
        Dim upper As String = $"Upper: {text.ToUpper()}"
        Console.WriteLine("Test 4: " & upper)
        If upper = "Upper: HELLO" Then
            Console.WriteLine("PASS: Method call interpolation")
        Else
            Console.WriteLine("FAIL: Method call interpolation, got: " & upper)
        End If

        ' Empty interpolation content
        Dim empty As String = $"No interpolation here"
        Console.WriteLine("Test 5: " & empty)
        If empty = "No interpolation here" Then
            Console.WriteLine("PASS: No interpolation")
        Else
            Console.WriteLine("FAIL: No interpolation, got: " & empty)
        End If

        ' Interpolation at start
        Dim age As Integer = 25
        Dim ageStr As String = $"{age} years old"
        Console.WriteLine("Test 6: " & ageStr)
        If ageStr = "25 years old" Then
            Console.WriteLine("PASS: Interpolation at start")
        Else
            Console.WriteLine("FAIL: Interpolation at start, got: " & ageStr)
        End If

        ' Interpolation at end
        Dim endStr As String = $"Age: {age}"
        Console.WriteLine("Test 7: " & endStr)
        If endStr = "Age: 25" Then
            Console.WriteLine("PASS: Interpolation at end")
        Else
            Console.WriteLine("FAIL: Interpolation at end, got: " & endStr)
        End If

        ' Adjacent interpolations
        Dim a As String = "AB"
        Dim b As String = "CD"
        Dim adjacent As String = $"{a}{b}"
        Console.WriteLine("Test 8: " & adjacent)
        If adjacent = "ABCD" Then
            Console.WriteLine("PASS: Adjacent interpolations")
        Else
            Console.WriteLine("FAIL: Adjacent interpolations, got: " & adjacent)
        End If

        Console.WriteLine("SUCCESS: Interpolated string tests complete")
    End Sub
End Module
