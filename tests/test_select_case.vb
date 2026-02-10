Module MainModule
    Sub Main()
        Console.WriteLine("Testing Select Case...")
        
        ' Test Exact Match
        Dim x As Integer = 2
        Select Case x
            Case 1
                Console.WriteLine("FAILURE: Matched 1")
            Case 2
                Console.WriteLine("SUCCESS: Matched 2")
            Case Else
                Console.WriteLine("FAILURE: Matched Else")
        End Select

        ' Test Range
        x = 5
        Select Case x
            Case 1 To 4
                Console.WriteLine("FAILURE: Matched 1-4")
            Case 5 To 10
                Console.WriteLine("SUCCESS: Matched 5-10")
            Case Else
                Console.WriteLine("FAILURE: Matched Else Range")
        End Select

        ' Test Is Comparison
        x = 20
        Select Case x
            Case Is < 10
                Console.WriteLine("FAILURE: < 10")
            Case Is > 15
                Console.WriteLine("SUCCESS: > 15")
            Case Else
                Console.WriteLine("FAILURE: Else Is")
        End Select

        ' Test Multiple Values
        x = 3
        Select Case x
            Case 1, 3, 5
                Console.WriteLine("SUCCESS: Matched Multiple")
            Case Else
                Console.WriteLine("FAILURE: Multiple Else")
        End Select

        ' Test Case Else
        x = 100
        Select Case x
            Case 1
                Console.WriteLine("FAILURE: 1")
            Case Else
                Console.WriteLine("SUCCESS: Matched Else Final")
        End Select

        ' Test Case String (Merged from root)
        Dim status As String = "Active"
        Select Case status
            Case "Pending"
                Console.WriteLine("FAILURE: Matched Pending")
            Case "Active"
                Console.WriteLine("SUCCESS: Matched Active String")
            Case Else
                Console.WriteLine("FAILURE: Matched Else String")
        End Select

        ' Test Exit Select (Merged from root)
        Dim val As Integer = 5
        Dim exitedCorrectly As Boolean = False
        Select Case val
            Case 1 To 10
                If val = 5 Then
                    exitedCorrectly = True
                    Exit Select
                    exitedCorrectly = False
                End If
            Case Else
                Console.WriteLine("FAILURE: Matched Else in Exit Test")
        End Select
        If exitedCorrectly Then
            Console.WriteLine("SUCCESS: Exit Select worked")
        Else
            Console.WriteLine("FAILURE: Exit Select failed")
        End If

        ' Test Nested Select (Merged from root)
        Dim category As Integer = 2
        Dim level As Integer = 3
        Select Case category
            Case 1
                Console.WriteLine("FAILURE: Category 1")
            Case 2
                Select Case level
                    Case 1, 2
                        Console.WriteLine("FAILURE: Level 1, 2")
                    Case 3
                        Console.WriteLine("SUCCESS: Nested Expert level")
                    Case Else
                        Console.WriteLine("FAILURE: Nested level Else")
                End Select
            Case Else
                Console.WriteLine("FAILURE: Category Else")
        End Select

        Console.WriteLine("Select Case Tests Completed")
    End Sub
End Module
