Module MainModule
    Sub Main()
        Dim passed = 0
        Dim failed = 0

        Console.WriteLine("=== Lambda Tests ===")

        ' -----------------------------------------------
        ' Test 1: Single-line Function lambda
        ' -----------------------------------------------
        Dim square = Function(x) x * x
        If square(5) = 25 Then
            Console.WriteLine("PASS: Single-line Function lambda")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Single-line Function lambda - got " & square(5))
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 2: Single-line Sub lambda
        ' -----------------------------------------------
        Dim subResult = ""
        Dim printer = Sub(msg) subResult = msg
        printer("Hello Lambda")
        If subResult = "Hello Lambda" Then
            Console.WriteLine("PASS: Single-line Sub lambda")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Single-line Sub lambda - got " & subResult)
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 3: Lambda with multiple parameters
        ' -----------------------------------------------
        Dim add = Function(a, b) a + b
        If add(3, 7) = 10 Then
            Console.WriteLine("PASS: Lambda with multiple parameters")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda with multiple parameters - got " & add(3, 7))
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 4: Lambda with no parameters
        ' -----------------------------------------------
        Dim getFortyTwo = Function() 42
        If getFortyTwo() = 42 Then
            Console.WriteLine("PASS: Lambda with no parameters")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda with no parameters - got " & getFortyTwo())
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 5: Closure capture
        ' -----------------------------------------------
        Dim factor = 10
        Dim multiplier = Function(x) x * factor
        If multiplier(5) = 50 Then
            Console.WriteLine("PASS: Closure capture")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Closure capture - got " & multiplier(5))
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 6: Lambda with string operations
        ' -----------------------------------------------
        Dim greet = Function(name) "Hello, " & name & "!"
        If greet("World") = "Hello, World!" Then
            Console.WriteLine("PASS: Lambda with string operations")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda with string operations - got " & greet("World"))
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 7: Lambda with boolean logic
        ' -----------------------------------------------
        Dim isEven = Function(n) (n Mod 2) = 0
        If isEven(4) = True And isEven(3) = False Then
            Console.WriteLine("PASS: Lambda with boolean logic")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda with boolean logic")
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 8: Nested lambda calls
        ' -----------------------------------------------
        Dim double = Function(x) x * 2
        Dim triple = Function(x) x * 3
        If double(triple(5)) = 30 Then
            Console.WriteLine("PASS: Nested lambda calls")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Nested lambda calls - got " & double(triple(5)))
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 9: Lambda returning another lambda (higher-order)
        ' -----------------------------------------------
        Dim makeAdder = Function(n) Function(x) x + n
        Dim add5 = makeAdder(5)
        If add5(10) = 15 Then
            Console.WriteLine("PASS: Lambda returning lambda (higher-order)")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda returning lambda - got " & add5(10))
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 10: Multi-line Function lambda
        ' -----------------------------------------------
        Dim clamp = Function(v, lo, hi)
            If v < lo Then
                Return lo
            ElseIf v > hi Then
                Return hi
            Else
                Return v
            End If
        End Function
        If clamp(-5, 0, 100) = 0 And clamp(50, 0, 100) = 50 And clamp(200, 0, 100) = 100 Then
            Console.WriteLine("PASS: Multi-line Function lambda")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Multi-line Function lambda")
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 11: Multi-line Sub lambda
        ' -----------------------------------------------
        Dim result1 = ""
        Dim appendLines = Sub(prefix)
            result1 = result1 & prefix & "A "
            result1 = result1 & prefix & "B"
        End Sub
        appendLines("X")
        If result1 = "XA XB" Then
            Console.WriteLine("PASS: Multi-line Sub lambda")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Multi-line Sub lambda - got '" & result1 & "'")
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 12: Lambda stored in array
        ' -----------------------------------------------
        Dim ops(2) As Object
        ops(0) = Function(a, b) a + b
        ops(1) = Function(a, b) a - b
        ops(2) = Function(a, b) a * b
        Dim addR = ops(0)
        Dim subR = ops(1)
        Dim mulR = ops(2)
        If addR(10, 3) = 13 And subR(10, 3) = 7 And mulR(10, 3) = 30 Then
            Console.WriteLine("PASS: Lambda stored in array")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda stored in array")
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 13: Sub lambda with Console.WriteLine
        ' -----------------------------------------------
        Dim logger = Sub(msg) Console.WriteLine("LOG: " & msg)
        logger("Test message")
        Console.WriteLine("PASS: Sub lambda with Console.WriteLine")
        passed = passed + 1

        ' -----------------------------------------------
        ' Test 14: Lambda composition
        ' -----------------------------------------------
        Dim inc = Function(x) x + 1
        Dim dbl = Function(x) x * 2
        ' compose: first double, then increment
        If inc(dbl(3)) = 7 Then
            Console.WriteLine("PASS: Lambda composition")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda composition - got " & inc(dbl(3)))
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Test 15: Lambda with conditional expression
        ' -----------------------------------------------
        Dim abs_val = Function(x) If(x >= 0, x, -x)
        If abs_val(5) = 5 And abs_val(-3) = 3 Then
            Console.WriteLine("PASS: Lambda with conditional expression")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: Lambda with conditional expression")
            failed = failed + 1
        End If

        ' -----------------------------------------------
        ' Summary
        ' -----------------------------------------------
        Console.WriteLine("")
        Console.WriteLine("=== Lambda Tests Complete ===")
        Console.WriteLine(passed & " passed, " & failed & " failed out of " & (passed + failed) & " tests")
    End Sub
End Module