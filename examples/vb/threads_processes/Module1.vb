Module MainModule
    Sub Main()
        Console.WriteLine("========================================")
        Console.WriteLine("  Threads & Processes Demo")
        Console.WriteLine("========================================")
        Console.WriteLine("")

        ' ── Thread.Sleep Demo ──────────────────────────────
        Console.WriteLine("--- Thread.Sleep ---")
        Console.WriteLine("Sleeping 500 ms ...")

        Dim sw As New System.Diagnostics.Stopwatch()
        sw.Start()
        System.Threading.Thread.Sleep(500)
        sw.Stop()

        Console.WriteLine("Woke up after " & sw.ElapsedMilliseconds & " ms")
        Console.WriteLine("")

        ' ── Countdown Timer ────────────────────────────────
        Console.WriteLine("--- Countdown Timer ---")
        Dim i As Integer
        For i = 3 To 1 Step -1
            Console.WriteLine("  " & i & " ...")
            System.Threading.Thread.Sleep(400)
        Next
        Console.WriteLine("  Go!")
        Console.WriteLine("")

        ' ── Progress Simulation ────────────────────────────
        Console.WriteLine("--- Progress Simulation ---")
        Dim pct As Integer
        For pct = 0 To 100 Step 20
            Dim bar As String = String.Format("[{0}%]", pct)
            Console.WriteLine("  " & bar)
            System.Threading.Thread.Sleep(200)
        Next
        Console.WriteLine("")

        ' ── Debug Output ───────────────────────────────────
        Console.WriteLine("--- Debug Output ---")
        Debug.WriteLine("This message goes to the debug output")
        Debug.Write("Debug: ")
        Debug.WriteLine("checking assertions ...")

        Dim value As Integer = 42
        Debug.Assert(value > 0, "Value must be positive")
        Console.WriteLine("Debug.Assert passed (value = " & value & ")")
        Console.WriteLine("")

        ' ── Process.Start Demo ─────────────────────────────
        Console.WriteLine("--- Process.Start ---")
        Console.WriteLine("Launching 'echo' via Process.Start ...")
        System.Diagnostics.Process.Start("echo", "Hello from a child process!")

        ' Give it a moment to run
        System.Threading.Thread.Sleep(300)
        Console.WriteLine("")

        ' ── Timed Work Simulation ──────────────────────────
        Console.WriteLine("--- Timed Work Simulation ---")
        Dim totalWork As Integer = 5
        Dim sw2 As New System.Diagnostics.Stopwatch()
        sw2.Start()

        Dim n As Integer
        For n = 1 To totalWork
            Console.WriteLine("  Processing task " & n & " of " & totalWork & " ...")
            ' Simulate variable work with sleep
            System.Threading.Thread.Sleep(150 * n)
        Next

        sw2.Stop()
        Console.WriteLine("  Total time: " & sw2.ElapsedMilliseconds & " ms")
        Console.WriteLine("")

        ' ── Multiple Stopwatches ───────────────────────────
        Console.WriteLine("--- Multiple Stopwatches ---")
        Dim swA As New System.Diagnostics.Stopwatch()
        Dim swB As New System.Diagnostics.Stopwatch()

        swA.Start()
        System.Threading.Thread.Sleep(200)
        swB.Start()
        System.Threading.Thread.Sleep(300)
        swA.Stop()
        swB.Stop()

        Console.WriteLine("  Stopwatch A: " & swA.ElapsedMilliseconds & " ms")
        Console.WriteLine("  Stopwatch B: " & swB.ElapsedMilliseconds & " ms")
        Console.WriteLine("  (A started before B, so A >= B)")
        Console.WriteLine("")

        Console.WriteLine("========================================")
        Console.WriteLine("  Demo Complete!")
        Console.WriteLine("========================================")
    End Sub
End Module
