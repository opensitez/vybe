Imports System.Threading.Tasks
Imports System.Threading
Imports System.Diagnostics

Module AsyncTest
    Sub Main()
        Console.WriteLine("--- Async & Threading Test ---")
        
        ' 1. Test Task.Run (Asynchronous)
        Console.WriteLine("[Task] Starting Task.Run...")
        Dim t = Task.Run(Function()
            Console.WriteLine("[Task Thread] Task is running in background...")
            Thread.Sleep(500)
            Console.WriteLine("[Task Thread] Task completed background work.")
            Return "Task Result Success"
        End Function)
        
        Console.WriteLine("[Main] Main thread is NOT blocked. Task isCompleted: " & t.IsCompleted)
        
        ' 2. Test Task.Delay (Asynchronous)
        Console.WriteLine("[Task] Starting Task.Delay...")
        Dim d = Task.Delay(300)
        Console.WriteLine("[Main] Delay started, isCompleted: " & d.IsCompleted)
        
        ' 3. Test Task.Wait and Task.Result (Blocking)
        Console.WriteLine("[Main] Waiting for Task Result...")
        Dim result = t.Result ' Should block until t is completed
        Console.WriteLine("[Main] Task finished! Result: " & result)
        Console.WriteLine("[Main] Task isCompleted: " & t.IsCompleted)
        
        ' 4. Test Thread.Start and Join
        Console.WriteLine("[Thread] Starting new thread...")
        Dim th = New Thread(Sub()
            Console.WriteLine("[Bg Thread] Background thread is running...")
            Thread.Sleep(400)
            Console.WriteLine("[Bg Thread] Background thread finishing.")
        End Sub)
        
        th.Start()
        Console.WriteLine("[Main] Thread IsAlive: " & th.IsAlive)
        
        Console.WriteLine("[Main] Joining thread...")
        th.Join() ' Should block until thread finishes
        Console.WriteLine("[Main] Thread Joined. IsAlive: " & th.IsAlive)
        
        ' 5. Test Process.Start and WaitForExit
        Console.WriteLine("[Process] Starting process (echo)...")
        Dim si = New ProcessStartInfo("/bin/echo", "Hello from Vybe Process")
        Dim p = Process.Start(si)
        Console.WriteLine("[Main] Process started, HasExited: " & p.HasExited)
        
        Console.WriteLine("[Main] Waiting for Process to exit...")
        p.WaitForExit()
        Console.WriteLine("[Main] Process Exited. ExitCode: " & p.ExitCode & ", HasExited: " & p.HasExited)
        
        Console.WriteLine("--- Async Test Finished ---")
    End Sub
End Module
