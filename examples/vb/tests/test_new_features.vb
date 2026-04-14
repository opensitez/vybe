Module Module1
    Sub Main()
        Console.WriteLine("=== Typed Exceptions ===")
        
        ' Test 1: Throw and Catch specific exception
        Try
            Throw New ArgumentException("Invalid argument value")
        Catch ex As ArgumentException
            Console.WriteLine("Caught ArgumentException: " & ex.Message)
        End Try

        ' Test 2: Division by zero
        Try
            Dim x As Integer = 1 / 0
        Catch ex As Exception
            Console.WriteLine("Caught: " & ex.Message)
        End Try

        ' Test 3: Throw generic exception
        Try
            Throw New InvalidOperationException("Cannot do that")
        Catch ex As Exception
            Console.WriteLine("Caught " & ex.Message)
        Finally
            Console.WriteLine("Finally block executed")
        End Try

        Console.WriteLine("")
        Console.WriteLine("=== Tuples ===")
        
        ' Test Tuple.Create
        Dim t = Tuple.Create("Hello", 42, True)
        Console.WriteLine("Tuple Item1: " & t.Item1)
        Console.WriteLine("Tuple Item2: " & t.Item2)
        Console.WriteLine("Tuple Item3: " & t.Item3)

        ' Test New Tuple
        Dim t2 = New Tuple("World", 99)
        Console.WriteLine("Tuple2 Item1: " & t2.Item1)
        Console.WriteLine("Tuple2 Item2: " & t2.Item2)

        Console.WriteLine("")
        Console.WriteLine("=== Nullable ===")
        
        Dim n1 = New Nullable(42)
        Console.WriteLine("HasValue: " & n1.HasValue)
        Console.WriteLine("Value: " & n1.Value)
        Console.WriteLine("GetValueOrDefault: " & n1.GetValueOrDefault(0))

        Console.WriteLine("")
        Console.WriteLine("=== Task / Async ===")
        
        ' Task.FromResult
        Dim task1 = Task.FromResult(42)
        Console.WriteLine("Task IsCompleted: " & task1.IsCompleted)
        Console.WriteLine("Task Result: " & task1.Result)
        
        ' Task.Delay (should just sleep briefly)
        Dim task2 = Task.Delay(10)
        Console.WriteLine("Task.Delay completed: " & task2.IsCompleted)

        ' Await
        Dim result = Await Task.FromResult("async result")
        Console.WriteLine("Await result: " & result)

        ' Task.Run with lambda
        Dim task3 = Task.Run(Function() 100 + 200)
        Console.WriteLine("Task.Run result: " & task3.Result)

        Console.WriteLine("")
        Console.WriteLine("=== BitConverter ===")
        Dim bytes = BitConverter.GetBytes(42)
        Console.WriteLine("42 as bytes: " & BitConverter.ToString(bytes))
        
        Dim backToInt = BitConverter.ToInt32(bytes, 0)
        Console.WriteLine("Back to int: " & backToInt)

        Console.WriteLine("")
        Console.WriteLine("=== MemoryStream ===")
        
        Dim ms = New MemoryStream()
        ms.WriteByte(72)
        ms.WriteByte(101)
        ms.WriteByte(108)
        ms.WriteByte(108)
        ms.WriteByte(111)
        Console.WriteLine("MemoryStream Length: " & ms.Length)
        Console.WriteLine("MemoryStream CanRead: " & ms.CanRead)

        Console.WriteLine("")
        Console.WriteLine("=== Mutex / Semaphore ===")
        
        Dim mtx = New Mutex(False)
        mtx.WaitOne()
        Console.WriteLine("Mutex acquired")
        mtx.ReleaseMutex()
        Console.WriteLine("Mutex released")

        Dim sem = New SemaphoreSlim(3, 10)
        Console.WriteLine("Semaphore count: " & sem.CurrentCount)
        sem.Wait()
        Console.WriteLine("After wait: " & sem.CurrentCount)
        Dim prev = sem.Release()
        Console.WriteLine("Previous count: " & prev)
        Console.WriteLine("After release: " & sem.CurrentCount)

        Console.WriteLine("")
        Console.WriteLine("All tests passed!")
    End Sub
End Module
