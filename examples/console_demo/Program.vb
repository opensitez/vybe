Module Program

    Sub Main()
        ' ── Title ────────────────────────────────────────────────
        Console.Title = "Irys Console Demo"

        ' ── Color Palette ────────────────────────────────────────
        Console.WriteLine("=== Console Color Palette ===")
        Console.WriteLine("")

        Dim colors() As String = { _
            "Black", "DarkBlue", "DarkGreen", "DarkCyan", _
            "DarkRed", "DarkMagenta", "DarkYellow", "Gray", _
            "DarkGray", "Blue", "Green", "Cyan", _
            "Red", "Magenta", "Yellow", "White" _
        }

        Dim i As Integer
        For i = 0 To 15
            Console.ForegroundColor = i
            Console.Write(colors(i) & "  ")
            If (i + 1) Mod 4 = 0 Then
                Console.WriteLine("")
            End If
        Next
        Console.ResetColor()
        Console.WriteLine("")

        ' ── Colored Output ───────────────────────────────────────
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("=== Colored Output Demo ===")

        Console.ForegroundColor = ConsoleColor.Green
        Console.Write("[OK] ")
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("Operation completed successfully")

        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write("[WARN] ")
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("Disk space is running low")

        Console.ForegroundColor = ConsoleColor.Red
        Console.Write("[ERR] ")
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("File not found: data.txt")

        Console.ResetColor()
        Console.WriteLine("")

        ' ── Background Colors ────────────────────────────────────
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("=== Background Colors ===")

        Console.ForegroundColor = ConsoleColor.White
        Console.BackgroundColor = ConsoleColor.DarkBlue
        Console.Write(" Info ")
        Console.ResetColor()
        Console.Write("  ")

        Console.ForegroundColor = ConsoleColor.White
        Console.BackgroundColor = ConsoleColor.DarkGreen
        Console.Write(" Success ")
        Console.ResetColor()
        Console.Write("  ")

        Console.ForegroundColor = ConsoleColor.Black
        Console.BackgroundColor = ConsoleColor.Yellow
        Console.Write(" Warning ")
        Console.ResetColor()
        Console.Write("  ")

        Console.ForegroundColor = ConsoleColor.White
        Console.BackgroundColor = ConsoleColor.DarkRed
        Console.Write(" Error ")
        Console.ResetColor()
        Console.WriteLine("")
        Console.WriteLine("")

        ' ── Interactive Input ────────────────────────────────────
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("=== Interactive Input ===")
        Console.ResetColor()

        Console.Write("What is your name? ")
        Dim name As String = Console.ReadLine()
        
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("Hello, " & name & "!")
        Console.ResetColor()
        Console.WriteLine("")

        ' ── Number Guessing Game ─────────────────────────────────
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("=== Quick Number Game ===")
        Console.ResetColor()
        Console.WriteLine("I'm thinking of a number between 1 and 10.")

        Dim secret As Integer = 7
        Dim guess As String
        Dim attempts As Integer = 0
        Dim found As Boolean = False

        Do While Not found
            Console.Write("Your guess: ")
            guess = Console.ReadLine()
            attempts = attempts + 1

            Dim guessNum As Integer = CInt(guess)
            If guessNum = secret Then
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Correct! You got it in " & attempts & " tries!")
                Console.ResetColor()
                found = True
            ElseIf guessNum < secret Then
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.WriteLine("Too low! Try again.")
                Console.ResetColor()
            Else
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.WriteLine("Too high! Try again.")
                Console.ResetColor()
            End If
        Loop

        Console.WriteLine("")

        ' ── Console Properties ───────────────────────────────────
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("=== Console Properties ===")
        Console.ResetColor()
        Console.WriteLine("Window Width:  " & Console.WindowWidth)
        Console.WriteLine("Window Height: " & Console.WindowHeight)
        Console.WriteLine("Buffer Width:  " & Console.BufferWidth)
        Console.WriteLine("Buffer Height: " & Console.BufferHeight)
        Console.WriteLine("")

        ' ── Goodbye ──────────────────────────────────────────────
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.WriteLine("Thanks for trying the Irys Console Demo!")
        Console.ResetColor()
    End Sub

End Module
