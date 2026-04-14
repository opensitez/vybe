Module MainModule
    Sub Main()
        Console.WriteLine("Testing Arrays...")

        ' Test 1: Array with size
        Dim numbers(5) As Integer
        numbers(0) = 10
        numbers(1) = 20
        numbers(2) = 30
        numbers(3) = 40
        numbers(4) = 50
        numbers(5) = 60

        If numbers(0) = 10 And numbers(5) = 60 Then
            Console.WriteLine("SUCCESS: Array indexed access")
        Else
            Console.WriteLine("FAILURE: Array indexed access")
        End If

        If UBound(numbers) = 5 And LBound(numbers) = 0 Then
            Console.WriteLine("SUCCESS: UBound and LBound")
        Else
            Console.WriteLine("FAILURE: UBound and LBound")
        End If

        ' Test 2: Array literal
        Dim fruits() As String = {"Apple", "Banana", "Cherry", "Date", "Elderberry"}

        If fruits(0) = "Apple" And fruits(2) = "Cherry" Then
            Console.WriteLine("SUCCESS: Array literal access")
        Else
            Console.WriteLine("FAILURE: Array literal access")
        End If

        If UBound(fruits) = 4 Then
            Console.WriteLine("SUCCESS: Array literal length")
        Else
            Console.WriteLine("FAILURE: Array literal length")
        End If

        ' Test 3: ReDim
        Dim items(3) As String
        items(0) = "First"
        items(1) = "Second"
        items(2) = "Third"
        items(3) = "Fourth"

        ' Resize preserving data
        ReDim Preserve items(7)
        If UBound(items) = 7 And items(0) = "First" Then
            Console.WriteLine("SUCCESS: ReDim Preserve")
        Else
            Console.WriteLine("FAILURE: ReDim Preserve")
        End If

        ' Add more items
        items(7) = "Eighth"
        If items(7) = "Eighth" Then
            Console.WriteLine("SUCCESS: Access after ReDim")
        Else
            Console.WriteLine("FAILURE: Access after ReDim")
        End If

        ' Test 4: Array with loop
        Dim scores(4) As Integer
        scores(0) = 80
        scores(1) = 90
        scores(2) = 70
        scores(3) = 95
        scores(4) = 85

        Dim total As Integer = 0
        Dim i As Integer
        For i = LBound(scores) To UBound(scores)
            total = total + scores(i)
        Next i

        If total = 420 Then
            Console.WriteLine("SUCCESS: Array loop total")
        Else
            Console.WriteLine("FAILURE: Array loop total. Got: " & CStr(total))
        End If

        Console.WriteLine("Array Tests Completed")
    End Sub
End Module
