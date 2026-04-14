Module MainModule
    Sub Main()
        Console.WriteLine("Testing Collections...")
        
        ' Test ArrayList
        Dim list As New ArrayList
        list.Add("A")
        list.Add("B")
        
        Console.WriteLine("Count: " & list.Count)
        If list.Count <> 2 Then
            Console.WriteLine("FAILURE: Count incorrect")
        End If
        
        Console.WriteLine("Item 0: " & list(0))
        If list(0) <> "A" Then
            Console.WriteLine("FAILURE: Item 0 incorrect")
        End If
        
        ' Test Generic List Proxy
        Dim genList As New List(Of String)
        genList.Add("X")
        genList.Add("Y")
        
        Console.WriteLine("Generic List Count: " & genList.Count)
        If genList.Count <> 2 Then
            Console.WriteLine("FAILURE: Generic List Count incorrect")
        End If
        
        ' Test Remove
        Call list.RemoveAt(0)
        Console.WriteLine("Count after Remove: " & list.Count)
        If list.Count <> 1 Then
             Console.WriteLine("FAILURE: Remove failed")
        End If
        
        ' Test Clear
        genList.Clear()
        Console.WriteLine("Generic List after Clear: " & genList.Count)
        
        Console.WriteLine("SUCCESS: Collections tests passed")
    End Sub
End Module
