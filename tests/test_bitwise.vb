Module BitwiseTest
    Sub Main()
        Console.WriteLine("Bitwise Test Start")

        ' Boolean Logic
        Dim t = True
        Dim f = False
        
        If t And t Then Console.WriteLine("True And True = True")
        If t Or f Then Console.WriteLine("True Or False = True")
        If t Xor f Then Console.WriteLine("True Xor False = True")
        If Not (t Xor t) Then Console.WriteLine("False Xor False = False") 
        ' (t Xor t is False, Not False is True)

        ' Bitwise Logic
        ' 5 = 101, 3 = 011
        ' 5 And 3 = 001 = 1
        ' 5 Or 3 = 111 = 7
        ' 5 Xor 3 = 110 = 6
        
        Dim a = 5
        Dim b = 3
        
        Console.WriteLine("5 And 3 = " & (a And b))
        Console.WriteLine("5 Or 3 = " & (a Or b))
        Console.WriteLine("5 Xor 3 = " & (a Xor b))
        
        ' Not 5 (assuming Long/Integer 32bit)
        ' Not 0...0101 = 1...1010 = -6
        Console.WriteLine("Not 5 = " & (Not a))

        ' Shifts
        ' 1 << 1 = 2
        ' 1 << 2 = 4
        ' 8 >> 1 = 4
        ' -8 >> 1 = -4 (Arithmetic shift)
        
        Console.WriteLine("1 << 1 = " & (1 << 1))
        Console.WriteLine("1 << 2 = " & (1 << 2))
        Console.WriteLine("8 >> 1 = " & (8 >> 1))
        Console.WriteLine("-8 >> 1 = " & (-8 >> 1))
        
        ' Precedence
        ' Or vs Xor
        ' True Or True Xor True
        ' If Xor > Or: True Or (True Xor True) -> True Or False -> True
        ' If Or > Xor: (True Or True) Xor True -> True Xor True -> False
        ' VB.NET: Xor > Or? No, docs say Not > And > Or > Xor.
        ' So (True Or True) Xor True -> True Xor True -> False.
        ' Wait, let's verify my grammar precedence.
        ' logical_xor = logical_or ~ (xor ~ logical_or)*
        ' So matches logical_or first.
        ' logical_xor parses "A Or B Xor C" as "A Or B" (logical_or) then Xor "C".
        ' So "(A Or B) Xor C".
        ' Meaning Or binds tighter than Xor in my parser?
        ' Yes. `logical_xor` wraps `logical_or`.
        ' So `logical_or` is evaluated "inside" `logical_xor`'s operands?
        ' No, `logical_xor` matches `logical_or` units.
        ' So `A Or B` is a valid `logical_or`.
        ' `A Or B Xor C`:
        ' `logical_xor` sees `logical_or` ("A Or B") then "Xor" then `logical_or` ("C").
        ' So it groups `(A Or B) Xor C`.
        ' This implies Or > Xor execution-wise (Or happens first).
        ' This matches standard VB6/VB.NET precedence where Xor is lowest logic op.
        
        If (True Or True Xor True) = False Then
            Console.WriteLine("Precedence: Or > Xor (Confirmed)")
        Else
            Console.WriteLine("Precedence: Xor > Or (Unexpected)")
        End If

        ' And vs Or
        ' True Or False And False
        ' If And > Or: True Or (False And False) -> True Or False -> True
        ' If Or > And: (True Or False) And False -> True And False -> False
        ' VB: And > Or.
        ' My grammar: logical_or wraps logical_and.
        ' "A Or B And C"
        ' logical_or sees `logical_and` ("A") Or `logical_and` ("B And C")?
        ' No, `logical_or = logical_and ~ (Or ~ logical_and)*`
        ' "A Or B And C":
        ' logical_and matches "A".
        ' Or matches "Or".
        ' logical_and matches "B And C" (because logical_and consumes the And chain).
        ' So it parses as `A Or (B And C)`.
        ' So And > Or. Confirmed.
        
        If (True Or False And False) = True Then
            Console.WriteLine("Precedence: And > Or (Confirmed)")
        Else
            Console.WriteLine("Precedence: Or > And (Unexpected)")
        End If

        Console.WriteLine("Bitwise Test Completed")
    End Sub
End Module
