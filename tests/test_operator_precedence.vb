Module OperatorPrecedenceTest
    Sub Main()
        Console.WriteLine("=== Operator Precedence Tests ===")

        ' -----------------------------------------------
        ' Integer Division (\)
        ' -----------------------------------------------
        Dim intDiv1 As Integer = 10 \ 3
        If intDiv1 = 3 Then
            Console.WriteLine("PASS: 10 \ 3 = 3")
        Else
            Console.WriteLine("FAIL: 10 \ 3 = " & intDiv1 & " (expected 3)")
        End If

        Dim intDiv2 As Integer = 7 \ 2
        If intDiv2 = 3 Then
            Console.WriteLine("PASS: 7 \ 2 = 3")
        Else
            Console.WriteLine("FAIL: 7 \ 2 = " & intDiv2 & " (expected 3)")
        End If

        Dim intDiv3 As Integer = 25 \ 7
        If intDiv3 = 3 Then
            Console.WriteLine("PASS: 25 \ 7 = 3")
        Else
            Console.WriteLine("FAIL: 25 \ 7 = " & intDiv3 & " (expected 3)")
        End If

        ' -----------------------------------------------
        ' Exponent (^)
        ' -----------------------------------------------
        Dim exp1 As Double = 2 ^ 10
        If exp1 = 1024 Then
            Console.WriteLine("PASS: 2 ^ 10 = 1024")
        Else
            Console.WriteLine("FAIL: 2 ^ 10 = " & exp1 & " (expected 1024)")
        End If

        Dim exp2 As Double = 3 ^ 3
        If exp2 = 27 Then
            Console.WriteLine("PASS: 3 ^ 3 = 27")
        Else
            Console.WriteLine("FAIL: 3 ^ 3 = " & exp2 & " (expected 27)")
        End If

        Dim exp3 As Double = 9 ^ 0.5
        If exp3 = 3 Then
            Console.WriteLine("PASS: 9 ^ 0.5 = 3")
        Else
            Console.WriteLine("FAIL: 9 ^ 0.5 = " & exp3 & " (expected 3)")
        End If

        ' -----------------------------------------------
        ' Precedence: ^ binds tighter than * and +
        ' 2 + 3 ^ 2 should be 2 + 9 = 11, not 25
        ' -----------------------------------------------
        Dim prec1 As Double = 2 + 3 ^ 2
        If prec1 = 11 Then
            Console.WriteLine("PASS: 2 + 3 ^ 2 = 11 (exponent before add)")
        Else
            Console.WriteLine("FAIL: 2 + 3 ^ 2 = " & prec1 & " (expected 11)")
        End If

        ' 4 * 2 ^ 3 should be 4 * 8 = 32, not 512
        Dim prec2 As Double = 4 * 2 ^ 3
        If prec2 = 32 Then
            Console.WriteLine("PASS: 4 * 2 ^ 3 = 32 (exponent before multiply)")
        Else
            Console.WriteLine("FAIL: 4 * 2 ^ 3 = " & prec2 & " (expected 32)")
        End If

        ' -----------------------------------------------
        ' Integer division precedence: same as * /
        ' 10 + 20 \ 3 should be 10 + 6 = 16
        ' -----------------------------------------------
        Dim prec3 As Integer = 10 + 20 \ 3
        If prec3 = 16 Then
            Console.WriteLine("PASS: 10 + 20 \ 3 = 16 (intdiv before add)")
        Else
            Console.WriteLine("FAIL: 10 + 20 \ 3 = " & prec3 & " (expected 16)")
        End If

        ' -----------------------------------------------
        ' Mod precedence
        ' -----------------------------------------------
        Dim mod1 As Integer = 17 Mod 5
        If mod1 = 2 Then
            Console.WriteLine("PASS: 17 Mod 5 = 2")
        Else
            Console.WriteLine("FAIL: 17 Mod 5 = " & mod1 & " (expected 2)")
        End If

        ' -----------------------------------------------
        ' AndAlso (short-circuit) vs And (bitwise/logical)
        ' -----------------------------------------------
        Dim scA As Boolean = True
        Dim scB As Boolean = False

        ' AndAlso should short-circuit: if left is False, right isn't evaluated
        Dim r1 As Boolean = False AndAlso True
        If r1 = False Then
            Console.WriteLine("PASS: False AndAlso True = False")
        Else
            Console.WriteLine("FAIL: False AndAlso True = " & r1)
        End If

        Dim r2 As Boolean = True AndAlso True
        If r2 = True Then
            Console.WriteLine("PASS: True AndAlso True = True")
        Else
            Console.WriteLine("FAIL: True AndAlso True = " & r2)
        End If

        ' And (bitwise on integers)
        Dim bitAnd As Integer = 12 And 10
        If bitAnd = 8 Then
            Console.WriteLine("PASS: 12 And 10 = 8 (bitwise)")
        Else
            Console.WriteLine("FAIL: 12 And 10 = " & bitAnd & " (expected 8)")
        End If

        ' -----------------------------------------------
        ' OrElse (short-circuit) vs Or (bitwise/logical)
        ' -----------------------------------------------
        Dim r3 As Boolean = True OrElse False
        If r3 = True Then
            Console.WriteLine("PASS: True OrElse False = True")
        Else
            Console.WriteLine("FAIL: True OrElse False = " & r3)
        End If

        Dim r4 As Boolean = False OrElse False
        If r4 = False Then
            Console.WriteLine("PASS: False OrElse False = False")
        Else
            Console.WriteLine("FAIL: False OrElse False = " & r4)
        End If

        ' Or (bitwise on integers)
        Dim bitOr As Integer = 12 Or 3
        If bitOr = 15 Then
            Console.WriteLine("PASS: 12 Or 3 = 15 (bitwise)")
        Else
            Console.WriteLine("FAIL: 12 Or 3 = " & bitOr & " (expected 15)")
        End If

        ' -----------------------------------------------
        ' Xor (bitwise)
        ' -----------------------------------------------
        Dim xorVal As Integer = 12 Xor 10
        If xorVal = 6 Then
            Console.WriteLine("PASS: 12 Xor 10 = 6")
        Else
            Console.WriteLine("FAIL: 12 Xor 10 = " & xorVal & " (expected 6)")
        End If

        ' -----------------------------------------------
        ' Not (bitwise and boolean)
        ' -----------------------------------------------
        Dim boolNot As Boolean = Not True
        If boolNot = False Then
            Console.WriteLine("PASS: Not True = False")
        Else
            Console.WriteLine("FAIL: Not True = " & boolNot)
        End If

        ' -----------------------------------------------
        ' Bit Shifts
        ' -----------------------------------------------
        Dim shl As Integer = 1 << 4
        If shl = 16 Then
            Console.WriteLine("PASS: 1 << 4 = 16")
        Else
            Console.WriteLine("FAIL: 1 << 4 = " & shl & " (expected 16)")
        End If

        Dim shr As Integer = 128 >> 3
        If shr = 16 Then
            Console.WriteLine("PASS: 128 >> 3 = 16")
        Else
            Console.WriteLine("FAIL: 128 >> 3 = " & shr & " (expected 16)")
        End If

        ' -----------------------------------------------
        ' String Concatenation (&)
        ' -----------------------------------------------
        Dim concat1 As String = "Hello" & " " & "World"
        If concat1 = "Hello World" Then
            Console.WriteLine("PASS: String concat with &")
        Else
            Console.WriteLine("FAIL: String concat = '" & concat1 & "'")
        End If

        ' -----------------------------------------------
        ' Negation
        ' -----------------------------------------------
        Dim neg1 As Integer = -5
        If neg1 = -5 Then
            Console.WriteLine("PASS: -5 = -5")
        Else
            Console.WriteLine("FAIL: -5 = " & neg1)
        End If

        Dim neg2 As Double = -(3.14)
        If neg2 < 0 Then
            Console.WriteLine("PASS: -(3.14) is negative")
        Else
            Console.WriteLine("FAIL: -(3.14) = " & neg2)
        End If

        ' -----------------------------------------------
        ' Mixed operator precedence
        ' 2 ^ 3 + 4 * 5 - 10 \ 3 Mod 2
        ' = 8 + 20 - 3 Mod 2  (Note: \ and Mod same level as * /)
        ' = 8 + 20 - 1
        ' = 27
        ' -----------------------------------------------
        Dim complex As Double = 2 ^ 3 + 4 * 5 - 10 \ 3 Mod 2
        Console.WriteLine("Complex expr: 2^3 + 4*5 - 10\3 Mod 2 = " & complex)

        ' -----------------------------------------------
        ' Comparison operators
        ' -----------------------------------------------
        If 5 > 3 Then Console.WriteLine("PASS: 5 > 3")
        If 3 < 5 Then Console.WriteLine("PASS: 3 < 5")
        If 5 >= 5 Then Console.WriteLine("PASS: 5 >= 5")
        If 5 <= 5 Then Console.WriteLine("PASS: 5 <= 5")
        If 3 <> 5 Then Console.WriteLine("PASS: 3 <> 5")
        If 5 = 5 Then Console.WriteLine("PASS: 5 = 5")

        Console.WriteLine("=== Operator Precedence Tests Done ===")
    End Sub
End Module
