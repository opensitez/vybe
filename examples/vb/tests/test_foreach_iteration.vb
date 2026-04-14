Imports System
Imports System.Collections

Module TestForEachIteration
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, name As String)
        If condition Then
            passed = passed + 1
            Console.WriteLine("  PASS: " & name)
        Else
            failed = failed + 1
            Console.WriteLine("  FAILURE: " & name)
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== ForEach Iteration Tests ===")

        ' --- Test 1: ForEach over Array ---
        Dim arr() As Integer = {10, 20, 30}
        Dim arrSum As Integer = 0
        For Each item In arr
            arrSum = arrSum + item
        Next
        Assert(arrSum = 60, "ForEach over Array sums correctly")

        ' --- Test 2: ForEach over ArrayList ---
        Dim al As New ArrayList()
        al.Add("apple")
        al.Add("banana")
        al.Add("cherry")
        Dim alResult As String = ""
        For Each fruit In al
            alResult = alResult & fruit & ","
        Next
        Assert(alResult = "apple,banana,cherry,", "ForEach over ArrayList")

        ' --- Test 3: ForEach over String (chars) ---
        Dim s As String = "ABC"
        Dim charResult As String = ""
        For Each ch In s
            charResult = charResult & ch & "-"
        Next
        Assert(charResult = "A-B-C-", "ForEach over String yields characters")

        ' --- Test 4: ForEach over Dictionary ---
        Dim dict As New Dictionary(Of String, Integer)
        dict.Add("x", 1)
        dict.Add("y", 2)
        dict.Add("z", 3)
        Dim dictKeys As String = ""
        Dim dictValSum As Integer = 0
        For Each kvp In dict
            dictKeys = dictKeys & kvp.Key & ","
            dictValSum = dictValSum + kvp.Value
        Next
        Assert(dictValSum = 6, "ForEach over Dictionary sums values")
        Assert(dictKeys = "x,y,z,", "ForEach over Dictionary iterates keys")

        ' --- Test 5: ForEach over Queue ---
        Dim q As New Queue()
        q.Enqueue("first")
        q.Enqueue("second")
        q.Enqueue("third")
        Dim qResult As String = ""
        For Each item In q
            qResult = qResult & item & ","
        Next
        Assert(qResult = "first,second,third,", "ForEach over Queue")

        ' --- Test 6: ForEach over Stack ---
        Dim stk As New Stack()
        stk.Push(1)
        stk.Push(2)
        stk.Push(3)
        Dim stkResult As String = ""
        For Each item In stk
            stkResult = stkResult & CStr(item) & ","
        Next
        ' Stack.ToArray returns top-first order
        Assert(stkResult = "3,2,1,", "ForEach over Stack (top-first)")

        ' --- Test 7: ForEach over HashSet ---
        Dim hs As New HashSet(Of String)
        hs.Add("red")
        hs.Add("green")
        hs.Add("blue")
        hs.Add("red") ' duplicate, should be ignored
        Dim hsCount As Integer = 0
        For Each color In hs
            hsCount = hsCount + 1
        Next
        Assert(hsCount = 3, "ForEach over HashSet (unique items only)")

        ' --- Test 8: Exit For inside ForEach ---
        Dim exitResult As String = ""
        For Each n In {1, 2, 3, 4, 5}
            If n = 3 Then Exit For
            exitResult = exitResult & CStr(n) & ","
        Next
        Assert(exitResult = "1,2,", "Exit For inside ForEach")

        ' --- Test 9: Continue For inside ForEach ---
        Dim contResult As String = ""
        For Each n In {1, 2, 3, 4, 5}
            If n = 3 Then Continue For
            contResult = contResult & CStr(n) & ","
        Next
        Assert(contResult = "1,2,4,5,", "Continue For inside ForEach")

        ' --- Test 10: ForEach over Nothing (empty) ---
        Dim emptyArr() As Integer = {}
        Dim emptyCount As Integer = 0
        For Each x In emptyArr
            emptyCount = emptyCount + 1
        Next
        Assert(emptyCount = 0, "ForEach over empty array")

        ' --- Summary ---
        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("SUCCESS: All ForEach iteration tests passed!")
        Else
            Console.WriteLine("FAILURE: Some ForEach tests failed")
        End If
    End Sub
End Module
