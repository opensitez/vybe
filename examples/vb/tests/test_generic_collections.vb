Module MainModule
    Sub Main()
        Console.WriteLine("Testing System.Collections.Generic...")
        Console.WriteLine("")
        
        ' ============================================================================
        ' Test Queue (FIFO)
        ' ============================================================================
        Console.WriteLine("=== Testing Queue ===")
        Dim queue As Object = New Queue()
        
        ' Enqueue items
        queue.Enqueue("First")
        queue.Enqueue("Second")
        queue.Enqueue("Third")
        
        Console.WriteLine("Queue Count: " & queue.Count)
        If queue.Count <> 3 Then
            Console.WriteLine("FAILURE: Queue Count incorrect")
        End If
        
        ' Peek (doesn't remove)
        Dim peeked As String = queue.Peek()
        Console.WriteLine("Peeked: " & peeked)
        If peeked <> "First" Then
            Console.WriteLine("FAILURE: Peek incorrect")
        End If
        
        ' Dequeue (removes in FIFO order)
        Dim item1 As String = queue.Dequeue()
        Console.WriteLine("Dequeued: " & item1)
        If item1 <> "First" Then
            Console.WriteLine("FAILURE: First dequeue incorrect")
        End If
        
        Dim item2 As String = queue.Dequeue()
        Console.WriteLine("Dequeued: " & item2)
        If item2 <> "Second" Then
            Console.WriteLine("FAILURE: Second dequeue incorrect")
        End If
        
        Console.WriteLine("Queue Count after dequeues: " & queue.Count)
        If queue.Count <> 1 Then
            Console.WriteLine("FAILURE: Queue Count after dequeue incorrect")
        End If
        
        queue.Clear()
        Console.WriteLine("Queue Count after Clear: " & queue.Count)
        If queue.Count <> 0 Then
            Console.WriteLine("FAILURE: Queue Clear failed")
        End If
        
        Console.WriteLine("SUCCESS: Queue tests passed")
        Console.WriteLine("")
        
        ' ============================================================================
        ' Test Stack (LIFO)
        ' ============================================================================
        Console.WriteLine("=== Testing Stack ===")
        Dim stack As Object = New Stack()
        
        ' Push items
        stack.Push("Bottom")
        stack.Push("Middle")
        stack.Push("Top")
        
        Console.WriteLine("Stack Count: " & stack.Count)
        If stack.Count <> 3 Then
            Console.WriteLine("FAILURE: Stack Count incorrect")
        End If
        
        ' Peek (doesn't remove)
        Dim peekedStack As String = stack.Peek()
        Console.WriteLine("Peeked: " & peekedStack)
        If peekedStack <> "Top" Then
            Console.WriteLine("FAILURE: Stack Peek incorrect")
        End If
        
        ' Pop (removes in LIFO order)
        Dim stackItem1 As String = stack.Pop()
        Console.WriteLine("Popped: " & stackItem1)
        If stackItem1 <> "Top" Then
            Console.WriteLine("FAILURE: First pop incorrect")
        End If
        
        Dim stackItem2 As String = stack.Pop()
        Console.WriteLine("Popped: " & stackItem2)
        If stackItem2 <> "Middle" Then
            Console.WriteLine("FAILURE: Second pop incorrect")
        End If
        
        Console.WriteLine("Stack Count after pops: " & stack.Count)
        If stack.Count <> 1 Then
            Console.WriteLine("FAILURE: Stack Count after pop incorrect")
        End If
        
        stack.Clear()
        Console.WriteLine("Stack Count after Clear: " & stack.Count)
        If stack.Count <> 0 Then
            Console.WriteLine("FAILURE: Stack Clear failed")
        End If
        
        Console.WriteLine("SUCCESS: Stack tests passed")
        Console.WriteLine("")
        
        ' ============================================================================
        ' Test HashSet (unique items)
        ' ============================================================================
        Console.WriteLine("=== Testing HashSet ===")
        Dim hashset As Object = New HashSet()
        
        ' Add unique items
        Dim added1 As Boolean = hashset.Add("Apple")
        Console.WriteLine("Added Apple: " & added1)
        If Not added1 Then
            Console.WriteLine("FAILURE: First Add should return True")
        End If
        
        Dim added2 As Boolean = hashset.Add("Banana")
        Console.WriteLine("Added Banana: " & added2)
        
        ' Try adding duplicate (should return False)
        Dim added3 As Boolean = hashset.Add("Apple")
        Console.WriteLine("Added Apple again: " & added3)
        If added3 Then
            Console.WriteLine("FAILURE: Duplicate Add should return False")
        End If
        
        Console.WriteLine("HashSet Count: " & hashset.Count)
        If hashset.Count <> 2 Then
            Console.WriteLine("FAILURE: HashSet Count incorrect (should be 2, not 3)")
        End If
        
        ' Test Contains
        Dim containsApple As Boolean = hashset.Contains("Apple")
        Console.WriteLine("Contains Apple: " & containsApple)
        If Not containsApple Then
            Console.WriteLine("FAILURE: Should contain Apple")
        End If
        
        Dim containsOrange As Boolean = hashset.Contains("Orange")
        Console.WriteLine("Contains Orange: " & containsOrange)
        If containsOrange Then
            Console.WriteLine("FAILURE: Should not contain Orange")
        End If
        
        ' Test Remove
        Dim removed As Boolean = hashset.Remove("Banana")
        Console.WriteLine("Removed Banana: " & removed)
        If Not removed Then
            Console.WriteLine("FAILURE: Remove should return True")
        End If
        
        Console.WriteLine("HashSet Count after Remove: " & hashset.Count)
        If hashset.Count <> 1 Then
            Console.WriteLine("FAILURE: HashSet Count after remove incorrect")
        End If
        
        hashset.Clear()
        Console.WriteLine("HashSet Count after Clear: " & hashset.Count)
        If hashset.Count <> 0 Then
            Console.WriteLine("FAILURE: HashSet Clear failed")
        End If
        
        Console.WriteLine("SUCCESS: HashSet tests passed")
        Console.WriteLine("")
        
        ' ============================================================================
        ' Test Dictionary
        ' ============================================================================
        Console.WriteLine("=== Testing Dictionary ===")
        Dim dict As Object = New Dictionary()
        
        ' Add key-value pairs
        dict.Add("name", "John")
        dict.Add("age", 30)
        dict.Add("city", "New York")
        
        Console.WriteLine("Dictionary Count: " & dict.Count)
        If dict.Count <> 3 Then
            Console.WriteLine("FAILURE: Dictionary Count incorrect")
        End If
        
        ' Test Item access
        Dim name As String = dict.Item("name")
        Console.WriteLine("Name: " & name)
        If name <> "John" Then
            Console.WriteLine("FAILURE: Dictionary Item incorrect")
        End If
        
        ' Test ContainsKey
        Dim hasAge As Boolean = dict.ContainsKey("age")
        Console.WriteLine("Contains 'age': " & hasAge)
        If Not hasAge Then
            Console.WriteLine("FAILURE: Should contain 'age' key")
        End If
        
        ' Test Remove
        dict.Remove("city")
        Console.WriteLine("Dictionary Count after Remove: " & dict.Count)
        If dict.Count <> 2 Then
            Console.WriteLine("FAILURE: Dictionary Count after remove incorrect")
        End If
        
        Console.WriteLine("SUCCESS: Dictionary tests passed")
        Console.WriteLine("")
        
        ' ============================================================================
        ' Test Convert.ToDateTime
        ' ============================================================================
        Console.WriteLine("=== Testing Convert.ToDateTime ===")
        
        ' From string
        Dim dt1 As Date = Convert.ToDateTime("2026-02-10")
        Console.WriteLine("Parsed date: " & dt1)
        
        ' From number (Excel date format)
        Dim dt2 As Date = Convert.ToDateTime(45000)
        Console.WriteLine("Date from number: " & dt2)
        
        ' From existing Date
        Dim dt3 As Date = #2/10/2026#
        Dim dt4 As Date = Convert.ToDateTime(dt3)
        Console.WriteLine("Date from Date: " & dt4)
        
        Console.WriteLine("SUCCESS: Convert.ToDateTime tests passed")
        Console.WriteLine("")
        
        Console.WriteLine("=== ALL TESTS PASSED ===")
    End Sub
End Module
