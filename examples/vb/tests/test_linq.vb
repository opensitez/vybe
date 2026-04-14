Imports System

Module TestLinqDebug2
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
        Console.WriteLine("=== LINQ Debug2 ===")

        Dim numbers() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}

        Console.WriteLine("Test: Select")
        Dim doubled = numbers.Select(Function(x) x * 2)
        Assert(doubled(0) = 2, "Select: first element doubled")
        Assert(doubled(4) = 10, "Select: fifth element doubled")
        Assert(doubled.Length = 10, "Select: preserves count")

        Console.WriteLine("Test: Where")
        Dim evens = numbers.Where(Function(x) x Mod 2 = 0)
        Assert(evens.Length = 5, "Where: 5 even numbers")
        Assert(evens(0) = 2, "Where: first even is 2")
        Assert(evens(4) = 10, "Where: last even is 10")

        Console.WriteLine("Test: First/FirstOrDefault")
        Assert(numbers.First() = 1, "First: returns first element")
        Assert(numbers.First(Function(x) x Mod 2 = 0) = 2, "First predicate")
        Assert(numbers.FirstOrDefault(Function(x) x > 100) = Nothing, "FirstOrDefault: Nothing")

        Console.WriteLine("Test: Last/LastOrDefault")
        Assert(numbers.Last() = 10, "Last: returns last")
        Assert(numbers.Last(Function(x) x Mod 2 = 0) = 10, "Last predicate")
        Assert(numbers.LastOrDefault(Function(x) x > 100) = Nothing, "LastOrDefault: Nothing")

        Console.WriteLine("Test: Count")
        Assert(numbers.Count() = 10, "Count: total")
        Assert(numbers.Count(Function(x) x Mod 2 = 0) = 5, "Count predicate")

        Console.WriteLine("Test: Any/All")
        Assert(numbers.Any(Function(x) x Mod 2 = 0) = True, "Any: has evens")
        Assert(numbers.All(Function(x) x > 0) = True, "All: positive")
        Assert(numbers.All(Function(x) x Mod 2 = 0) = False, "All: not all even")

        Console.WriteLine("Test: Sum/Min/Max/Average")
        Assert(numbers.Sum() = 55, "Sum: 55")
        Assert(numbers.Min() = 1, "Min: 1")
        Assert(numbers.Max() = 10, "Max: 10")
        Assert(numbers.Average() = 5.5, "Average: 5.5")

        Console.WriteLine("Test: OrderBy")
        Dim unordered() As Integer = {3, 1, 4, 1, 5, 9}
        Dim ordered = unordered.OrderBy(Function(x) x)
        Assert(ordered(0) = 1, "OrderBy: first is smallest")
        Assert(ordered(5) = 9, "OrderBy: last is largest")
        Dim descOrdered = unordered.OrderByDescending(Function(x) x)
        Assert(descOrdered(0) = 9, "OrderByDescending: first is largest")

        Console.WriteLine("Test: Skip/Take")
        Dim skipped = numbers.Skip(7)
        Assert(skipped.Length = 3, "Skip: 3 remaining")
        Assert(skipped(0) = 8, "Skip: first remaining is 8")
        Dim taken = numbers.Take(3)
        Assert(taken.Length = 3, "Take: 3 elements")
        Assert(taken(2) = 3, "Take: third is 3")

        Console.WriteLine("Test: Distinct")
        Dim dupes() As Integer = {1, 2, 2, 3, 3, 3}
        Dim unique = dupes.Distinct()
        Assert(unique.Length = 3, "Distinct: 3 unique")

        Console.WriteLine("Test: Chaining")
        Dim chainResult = numbers.Where(Function(x) x > 5).Select(Function(x) x * 10)
        Assert(chainResult.Length = 5, "Chain: 5 results")
        Assert(chainResult(0) = 60, "Chain: first is 60")

        Console.WriteLine("Test: ToList")
        Dim asList = numbers.ToList()
        Assert(asList.Count = 10, "ToList: 10 items")

        Console.WriteLine("Test: Reverse")
        Dim rev = numbers.Take(3).Reverse()
        Assert(rev(0) = 3, "Reverse: first becomes last")
        Assert(rev(2) = 1, "Reverse: last becomes first")

        Console.WriteLine("Test: Contains")
        Assert(numbers.Contains(5) = True, "Contains: 5 present")
        Assert(numbers.Contains(99) = False, "Contains: 99 absent")

        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("SUCCESS: All LINQ tests passed!")
        Else
            Console.WriteLine("FAILURE: Some LINQ tests failed")
        End If
    End Sub
End Module
