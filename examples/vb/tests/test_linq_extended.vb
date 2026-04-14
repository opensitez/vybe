Imports System

Module TestLinqExtended
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, name As String)
        If condition Then
            passed = passed + 1
            Console.WriteLine("  PASS: " & name)
        Else
            failed = failed + 1
            Console.WriteLine("  FAIL: " & name)
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== LINQ Extended Operators ===")

        Dim numbers() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}

        ' --- GroupBy ---
        Console.WriteLine("Test: GroupBy")
        Dim grouped = numbers.GroupBy(Function(x) If(x Mod 2 = 0, "even", "odd"))
        Assert(grouped.Length = 2, "GroupBy: 2 groups")
        ' Groups should have Key and Items
        Dim g1 = grouped(0)
        Assert(g1.Key = "odd", "GroupBy: first group key is odd")
        Assert(g1.Count = 5, "GroupBy: odd group has 5 items")

        ' --- Union ---
        Console.WriteLine("Test: Union")
        Dim a() As Integer = {1, 2, 3, 4}
        Dim b() As Integer = {3, 4, 5, 6}
        Dim unioned = a.Union(b)
        Assert(unioned.Length = 6, "Union: 6 unique elements")
        Assert(unioned(0) = 1, "Union: starts with 1")
        Assert(unioned(5) = 6, "Union: ends with 6")

        ' --- Intersect ---
        Console.WriteLine("Test: Intersect")
        Dim inter = a.Intersect(b)
        Assert(inter.Length = 2, "Intersect: 2 common elements")
        Assert(inter(0) = 3, "Intersect: first common is 3")
        Assert(inter(1) = 4, "Intersect: second common is 4")

        ' --- Except ---
        Console.WriteLine("Test: Except")
        Dim diff = a.Except(b)
        Assert(diff.Length = 2, "Except: 2 unique to a")
        Assert(diff(0) = 1, "Except: first is 1")
        Assert(diff(1) = 2, "Except: second is 2")

        ' --- Concat ---
        Console.WriteLine("Test: Concat")
        Dim c1() As Integer = {1, 2}
        Dim c2() As Integer = {3, 4}
        Dim concatenated = c1.Concat(c2)
        Assert(concatenated.Length = 4, "Concat: 4 total elements")
        Assert(concatenated(0) = 1, "Concat: first is 1")
        Assert(concatenated(3) = 4, "Concat: last is 4")

        ' --- SkipWhile ---
        Console.WriteLine("Test: SkipWhile")
        Dim sw = numbers.SkipWhile(Function(x) x < 5)
        Assert(sw.Length = 6, "SkipWhile: 6 elements remaining")
        Assert(sw(0) = 5, "SkipWhile: first is 5")

        ' --- TakeWhile ---
        Console.WriteLine("Test: TakeWhile")
        Dim tw = numbers.TakeWhile(Function(x) x < 5)
        Assert(tw.Length = 4, "TakeWhile: 4 elements taken")
        Assert(tw(3) = 4, "TakeWhile: last taken is 4")

        ' --- ElementAt ---
        Console.WriteLine("Test: ElementAt")
        Assert(numbers.ElementAt(0) = 1, "ElementAt: index 0 is 1")
        Assert(numbers.ElementAt(9) = 10, "ElementAt: index 9 is 10")

        ' --- ElementAtOrDefault ---
        Console.WriteLine("Test: ElementAtOrDefault")
        Assert(numbers.ElementAtOrDefault(0) = 1, "ElementAtOrDefault: index 0")
        Assert(numbers.ElementAtOrDefault(99) = Nothing, "ElementAtOrDefault: out of range")

        ' --- DefaultIfEmpty ---
        Console.WriteLine("Test: DefaultIfEmpty")
        Dim emptyArr() As Integer = {}
        Dim defaulted = emptyArr.DefaultIfEmpty(42)
        Assert(defaulted.Length = 1, "DefaultIfEmpty: 1 element for empty")
        Assert(defaulted(0) = 42, "DefaultIfEmpty: default value is 42")
        Dim nonEmpty = numbers.DefaultIfEmpty(42)
        Assert(nonEmpty.Length = 10, "DefaultIfEmpty: non-empty unchanged")

        ' --- SequenceEqual ---
        Console.WriteLine("Test: SequenceEqual")
        Dim s1() As Integer = {1, 2, 3}
        Dim s2() As Integer = {1, 2, 3}
        Dim s3() As Integer = {1, 2, 4}
        Assert(s1.SequenceEqual(s2) = True, "SequenceEqual: identical sequences")
        Assert(s1.SequenceEqual(s3) = False, "SequenceEqual: different sequences")

        ' --- ThenBy ---
        Console.WriteLine("Test: ThenBy")
        Dim sorted = numbers.OrderBy(Function(x) x)
        Dim thenSorted = sorted.ThenBy(Function(x) x)
        Assert(thenSorted(0) = 1, "ThenBy: first still 1")
        Assert(thenSorted(9) = 10, "ThenBy: last still 10")

        ' --- SelectMany ---
        Console.WriteLine("Test: SelectMany")
        Dim nested() As Object = {{1, 2}, {3, 4}, {5, 6}}
        Dim flat = nested.SelectMany(Function(x) x)
        Assert(flat.Length = 6, "SelectMany: flattened to 6")
        Assert(flat(0) = 1, "SelectMany: first is 1")
        Assert(flat(5) = 6, "SelectMany: last is 6")

        ' --- Aggregate ---
        Console.WriteLine("Test: Aggregate")
        Dim sum = numbers.Aggregate(0, Function(acc, x) acc + x)
        Assert(sum = 55, "Aggregate with seed: sum is 55")
        Dim product = numbers.Take(5).Aggregate(Function(acc, x) acc * x)
        Assert(product = 120, "Aggregate without seed: 5! = 120")

        ' --- Chaining new operators ---
        Console.WriteLine("Test: Chaining new operators")
        Dim chain = numbers.Where(Function(x) x > 3).SkipWhile(Function(x) x < 6).Take(3)
        Assert(chain.Length = 3, "Chain: 3 results")
        Assert(chain(0) = 6, "Chain: first is 6")
        Assert(chain(2) = 8, "Chain: last is 8")

        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("=== ALL LINQ EXTENDED TESTS PASSED ===")
        Else
            Console.WriteLine("=== SOME LINQ EXTENDED TESTS FAILED ===")
        End If
    End Sub
End Module
