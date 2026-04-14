Console.WriteLine("=== Standalone Test ===")

Dim x As Integer = 42
Console.WriteLine("x = " & x)

Dim queue As Object = New Queue()
queue.Enqueue("First")
queue.Enqueue("Second")

Console.WriteLine("Queue count: " & queue.Count)

Dim item As String = queue.Dequeue()
Console.WriteLine("Dequeued: " & item)

Console.WriteLine("=== TEST PASSED ===")
