Module CollectionsDemo
    Sub Main()
        Console.WriteLine("=== Collections Demo ===")
        Console.WriteLine("")
        
        ' Demonstrate Queue for task processing
        DemoTaskQueue()
        Console.WriteLine("")
        
        ' Demonstrate Stack for undo functionality
        DemoUndoStack()
        Console.WriteLine("")
        
        ' Demonstrate HashSet for unique values
        DemoUniqueVisitors()
        Console.WriteLine("")
        
        ' Demonstrate Dictionary for data storage
        DemoUserDatabase()
        Console.WriteLine("")
        
        ' Demonstrate DateTime conversions
        DemoDateTimeConversions()
    End Sub
    
    ' ============================================================================
    ' Queue Demo: Task Processing System
    ' ============================================================================
    Sub DemoTaskQueue()
        Console.WriteLine("--- Queue Demo: Task Queue System ---")
        
        Dim taskQueue As Object = New Queue()
        
        ' Add tasks to the queue
        taskQueue.Enqueue("Send email to customer")
        taskQueue.Enqueue("Process payment")
        taskQueue.Enqueue("Update inventory")
        taskQueue.Enqueue("Generate report")
        
        Console.WriteLine("Tasks in queue: " & taskQueue.Count)
        Console.WriteLine("")
        
        ' Process tasks in FIFO order
        Console.WriteLine("Processing tasks...")
        Dim taskNum As Integer = 1
        Do While taskQueue.Count > 0
            Dim task As String = taskQueue.Dequeue()
            Console.WriteLine("  " & taskNum & ". " & task)
            taskNum = taskNum + 1
        Loop
        
        Console.WriteLine("All tasks completed!")
    End Sub
    
    ' ============================================================================
    ' Stack Demo: Undo/Redo Functionality
    ' ============================================================================
    Sub DemoUndoStack()
        Console.WriteLine("--- Stack Demo: Text Editor Undo System ---")
        
        Dim undoStack As Object = New Stack()
        Dim text As String = ""
        
        ' Perform operations and save to undo stack
        text = "Hello"
        undoStack.Push(text)
        Console.WriteLine("Typed: " & text)
        
        text = text & " World"
        undoStack.Push(text)
        Console.WriteLine("Typed: " & text)
        
        text = text & "!"
        undoStack.Push(text)
        Console.WriteLine("Typed: " & text)
        
        Console.WriteLine("")
        Console.WriteLine("Current text: '" & text & "'")
        Console.WriteLine("History depth: " & undoStack.Count)
        Console.WriteLine("")
        
        ' Undo operations (LIFO order)
        Console.WriteLine("Pressing Undo...")
        text = undoStack.Pop()
        Console.WriteLine("After first undo: '" & text & "'")
        
        Console.WriteLine("Pressing Undo again...")
        text = undoStack.Pop()
        Console.WriteLine("After second undo: '" & text & "'")
    End Sub
    
    ' ============================================================================
    ' HashSet Demo: Tracking Unique Visitors
    ' ============================================================================
    Sub DemoUniqueVisitors()
        Console.WriteLine("--- HashSet Demo: Website Visitor Tracking ---")
        
        Dim visitors As Object = New HashSet()
        
        ' Simulate visitor tracking (some repeat visits)
        Dim visits As String() = {"user123", "user456", "user123", "user789", "user456", "user123", "user999"}
        
        Console.WriteLine("Processing " & visits.Length & " page visits...")
        
        Dim i As Integer
        For i = 0 To visits.Length - 1
            Dim wasNew As Boolean = visitors.Add(visits(i))
            If wasNew Then
                Console.WriteLine("  New visitor: " & visits(i))
            Else
                Console.WriteLine("  Returning visitor: " & visits(i))
            End If
        Next
        
        Console.WriteLine("")
        Console.WriteLine("Total unique visitors: " & visitors.Count)
        Console.WriteLine("Total page views: " & visits.Length)
    End Sub
    
    ' ============================================================================
    ' Dictionary Demo: User Database
    ' ============================================================================
    Sub DemoUserDatabase()
        Console.WriteLine("--- Dictionary Demo: User Profile System ---")
        
        ' Create user profiles
        Dim user1 As Object = New Dictionary()
        user1.Add("id", 101)
        user1.Add("name", "Alice Johnson")
        user1.Add("email", "alice@example.com")
        user1.Add("role", "Admin")
        
        Dim user2 As Object = New Dictionary()
        user2.Add("id", 102)
        user2.Add("name", "Bob Smith")
        user2.Add("email", "bob@example.com")
        user2.Add("role", "User")
        
        ' Store users in a dictionary by ID
        Dim userDatabase As Object = New Dictionary()
        userDatabase.Add(101, user1)
        userDatabase.Add(102, user2)
        
        Console.WriteLine("Users in database: " & userDatabase.Count)
        Console.WriteLine("")
        
        ' Lookup user by ID
        Dim userId As Integer = 101
        If userDatabase.ContainsKey(userId) Then
            Dim user As Object = userDatabase.Item(userId)
            Console.WriteLine("User #" & userId & " found:")
            Console.WriteLine("  Name: " & user.Item("name"))
            Console.WriteLine("  Email: " & user.Item("email"))
            Console.WriteLine("  Role: " & user.Item("role"))
        Else
            Console.WriteLine("User #" & userId & " not found")
        End If
    End Sub
    
    ' ============================================================================
    ' DateTime Conversion Demo
    ' ============================================================================
    Sub DemoDateTimeConversions()
        Console.WriteLine("--- DateTime Conversion Demo ---")
        
        ' Parse from string
        Dim dateStr As String = "2026-02-10"
        Dim parsedDate As Date = Convert.ToDateTime(dateStr)
        Console.WriteLine("Parsed '" & dateStr & "' as: " & parsedDate)
        
        ' Convert from different formats
        Dim dateStr2 As String = "02/10/2026"
        Dim parsedDate2 As Date = Convert.ToDateTime(dateStr2)
        Console.WriteLine("Parsed '" & dateStr2 & "' as: " & parsedDate2)
        
        ' Convert from numeric value (Excel date format)
        Dim excelDate As Double = 45000
        Dim dateFromNumber As Date = Convert.ToDateTime(excelDate)
        Console.WriteLine("Excel date " & excelDate & " as: " & dateFromNumber)
        
        ' Get current info
        Dim ts As Object = TimeSpan.FromDays(7)
        Console.WriteLine("")
        Console.WriteLine("7 days = " & ts.TotalHours & " hours")
        Console.WriteLine("7 days = " & ts.TotalMinutes & " minutes")
        Console.WriteLine("7 days = " & ts.TotalSeconds & " seconds")
    End Sub
End Module
