Imports System.Collections.Concurrent

Sub Main()
    Console.WriteLine("Testing ConcurrentDictionary...")
    Dim d As New ConcurrentDictionary
    d.TryAdd("key1", "value1")
    
    Dim v As Variant
    v = ""
    Dim result As Boolean
    result = d.TryGetValue("key1", v)
    Console.WriteLine("TryGetValue Result: " & result)
    Console.WriteLine("Value: " & v)
    
    d.AddOrUpdate("key2", "initial", "updated")
    Console.WriteLine("AddOrUpdate (Add): " & d.GetOrAdd("key2", "exists"))
    
    d.AddOrUpdate("key2", "initial", "final")
    Console.WriteLine("AddOrUpdate (Update): " & d.GetOrAdd("key2", "exists"))
    
    Console.WriteLine("Testing ConcurrentQueue...")
    Dim q As New ConcurrentQueue
    q.Enqueue(100)
    q.Enqueue(200)
    
    Dim i As Variant
    q.TryDequeue(i)
    Console.WriteLine("Dequeued: " & i)
    q.TryPeek(i)
    Console.WriteLine("Peek: " & i)
    Console.WriteLine("Count: " & q.Count)
    
    Console.WriteLine("Testing ConcurrentStack...")
    Dim s As New ConcurrentStack
    s.Push(10)
    s.Push(20)
    
    Dim j As Variant
    s.TryPop(j)
    Console.WriteLine("Popped: " & j)
    s.TryPeek(j)
    Console.WriteLine("Peek: " & j)
    Console.WriteLine("Count: " & s.Count)
End Sub
