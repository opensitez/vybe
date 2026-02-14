Imports System
Imports System.Collections
Module TestArrayListComplete
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, msg As String)
        If condition Then
            passed += 1
        Else
            Console.WriteLine("FAIL: " & msg)
            failed += 1
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== ArrayList/Collection Complete API Tests ===")
        Console.WriteLine()

        ' ===== Core methods (sanity check) =====
        Dim al As New ArrayList()
        al.Add("A")
        al.Add("B")
        al.Add("C")
        al.Add("D")
        al.Add("E")
        Assert(al.Count = 5, "Count = 5 after 5 adds")
        Assert(al.Item(0) = "A", "Item(0) = A")
        Assert(al.Item(4) = "E", "Item(4) = E")

        ' ===== Capacity =====
        Assert(al.Capacity >= al.Count, "Capacity >= Count")

        ' ===== Contains =====
        Assert(al.Contains("C") = True, "Contains C")
        Assert(al.Contains("Z") = False, "Not Contains Z")

        ' ===== IndexOf overloads =====
        al.Add("B") ' Now: A,B,C,D,E,B
        Assert(al.IndexOf("B") = 1, "IndexOf B = 1")
        Assert(al.IndexOf("B", 2) = 5, "IndexOf B from 2 = 5")
        Assert(al.IndexOf("B", 2, 2) = -1, "IndexOf B from 2 count 2 = -1 (only C,D)")
        Assert(al.IndexOf("B", 2, 4) = 5, "IndexOf B from 2 count 4 = 5")
        al.RemoveAt(5) ' Back to A,B,C,D,E

        ' ===== LastIndexOf overloads =====
        al.Add("B") ' A,B,C,D,E,B
        Assert(al.LastIndexOf("B") = 5, "LastIndexOf B = 5")
        Assert(al.LastIndexOf("B", 3) = 1, "LastIndexOf B ending at 3 = 1")
        al.RemoveAt(5) ' Back to A,B,C,D,E

        ' ===== Insert =====
        al.Insert(2, "X")  ' A,B,X,C,D,E
        Assert(al.Count = 6, "Insert increases count")
        Assert(al.Item(2) = "X", "Insert at 2 = X")
        Assert(al.Item(3) = "C", "After insert, C shifted to 3")
        al.RemoveAt(2) ' Back to A,B,C,D,E

        ' ===== AddRange =====
        Dim extra As New ArrayList()
        extra.Add("F")
        extra.Add("G")
        al.AddRange(extra)  ' A,B,C,D,E,F,G
        Assert(al.Count = 7, "AddRange adds 2 items")
        Assert(al.Item(5) = "F", "AddRange item F")
        Assert(al.Item(6) = "G", "AddRange item G")

        ' ===== InsertRange =====
        Dim ins As New ArrayList()
        ins.Add("Y")
        ins.Add("Z")
        al.InsertRange(2, ins)  ' A,B,Y,Z,C,D,E,F,G
        Assert(al.Count = 9, "InsertRange increases count by 2")
        Assert(al.Item(2) = "Y", "InsertRange Y at 2")
        Assert(al.Item(3) = "Z", "InsertRange Z at 3")
        Assert(al.Item(4) = "C", "After InsertRange, C at 4")

        ' ===== RemoveRange =====
        al.RemoveRange(2, 2)  ' Remove Y,Z -> A,B,C,D,E,F,G
        Assert(al.Count = 7, "RemoveRange removes 2")
        Assert(al.Item(2) = "C", "After RemoveRange, C back at 2")

        ' ===== GetRange =====
        Dim sub1 As ArrayList = al.GetRange(1, 3)  ' B,C,D
        Assert(sub1.Count = 3, "GetRange count = 3")
        Assert(sub1.Item(0) = "B", "GetRange(0) = B")
        Assert(sub1.Item(2) = "D", "GetRange(2) = D")

        ' ===== SetRange =====
        Dim rep As New ArrayList()
        rep.Add("P")
        rep.Add("Q")
        al.SetRange(1, rep)  ' A,P,Q,D,E,F,G
        Assert(al.Item(1) = "P", "SetRange(1) = P")
        Assert(al.Item(2) = "Q", "SetRange(2) = Q")
        Assert(al.Item(3) = "D", "SetRange doesn't shift D")

        ' ===== Reverse =====
        Dim rv As New ArrayList()
        rv.Add(1)
        rv.Add(2)
        rv.Add(3)
        rv.Add(4)
        rv.Add(5)
        rv.Reverse()
        Assert(rv.Item(0) = 5, "Reverse: first = 5")
        Assert(rv.Item(4) = 1, "Reverse: last = 1")

        ' ===== Reverse(index, count) =====
        Dim rv2 As New ArrayList()
        rv2.Add(1)
        rv2.Add(2)
        rv2.Add(3)
        rv2.Add(4)
        rv2.Add(5)
        rv2.Reverse(1, 3)  ' 1, [4,3,2], 5
        Assert(rv2.Item(0) = 1, "Reverse(1,3): first unchanged")
        Assert(rv2.Item(1) = 4, "Reverse(1,3): 4")
        Assert(rv2.Item(2) = 3, "Reverse(1,3): 3")
        Assert(rv2.Item(3) = 2, "Reverse(1,3): 2")
        Assert(rv2.Item(4) = 5, "Reverse(1,3): last unchanged")

        ' ===== Sort =====
        Dim sorted As New ArrayList()
        sorted.Add("Banana")
        sorted.Add("Apple")
        sorted.Add("Cherry")
        sorted.Sort()
        Assert(sorted.Item(0) = "Apple", "Sort: Apple first")
        Assert(sorted.Item(1) = "Banana", "Sort: Banana second")
        Assert(sorted.Item(2) = "Cherry", "Sort: Cherry third")

        ' ===== BinarySearch (on sorted list) =====
        Dim idx As Integer = sorted.BinarySearch("Banana")
        Assert(idx = 1, "BinarySearch Banana = 1")

        ' ===== Clone =====
        Dim cloned As ArrayList = al.Clone()
        Assert(cloned.Count = al.Count, "Clone has same count")
        Assert(cloned.Item(0) = al.Item(0), "Clone item 0 matches")
        cloned.Add("NEW")
        Assert(cloned.Count = al.Count + 1, "Clone is independent")

        ' ===== ToArray =====
        Dim arrList As New ArrayList()
        arrList.Add(10)
        arrList.Add(20)
        arrList.Add(30)
        Dim arr() As Object = arrList.ToArray()
        Assert(arr.Length = 3, "ToArray length = 3")
        Assert(arr(0) = 10, "ToArray(0) = 10")

        ' ===== CopyTo =====
        Dim copied() As Object = arrList.CopyTo()
        Assert(copied.Length = 3, "CopyTo length = 3")

        ' ===== TrimToSize =====
        arrList.TrimToSize()
        Assert(arrList.Count = 3, "TrimToSize preserves count")

        ' ===== Properties =====
        Assert(arrList.IsFixedSize = False, "IsFixedSize = False")
        Assert(arrList.IsReadOnly = False, "IsReadOnly = False")
        Assert(arrList.IsSynchronized = False, "IsSynchronized = False")

        ' ===== FindIndex =====
        Dim nums As New ArrayList()
        nums.Add(10)
        nums.Add(20)
        nums.Add(30)
        nums.Add(40)
        nums.Add(50)
        Dim fi As Integer = nums.FindIndex(Function(x) x > 25)
        Assert(fi = 2, "FindIndex x > 25 = 2 (value 30)")

        ' ===== FindLast =====
        Dim fl As Object = nums.FindLast(Function(x) x < 35)
        Assert(fl = 30, "FindLast x < 35 = 30")

        ' ===== FindLastIndex =====
        Dim fli As Integer = nums.FindLastIndex(Function(x) x < 35)
        Assert(fli = 2, "FindLastIndex x < 35 = 2")

        ' ===== TrueForAll =====
        Assert(nums.TrueForAll(Function(x) x > 0) = True, "TrueForAll > 0")
        Assert(nums.TrueForAll(Function(x) x > 15) = False, "TrueForAll > 15 = False")

        ' ===== ConvertAll =====
        Dim doubled As ArrayList = nums.ConvertAll(Function(x) x * 2)
        Assert(doubled.Count = 5, "ConvertAll count = 5")
        Assert(doubled.Item(0) = 20, "ConvertAll(0) = 20")
        Assert(doubled.Item(4) = 100, "ConvertAll(4) = 100")

        ' ===== Find, FindAll, Exists, RemoveAll =====
        Dim found As Object = nums.Find(Function(x) x > 25)
        Assert(found = 30, "Find x > 25 = 30")

        Dim all As Object = nums.FindAll(Function(x) x >= 30)
        Assert(all.Length = 3, "FindAll >= 30 has 3 items")

        Assert(nums.Exists(Function(x) x = 40) = True, "Exists 40")
        Assert(nums.Exists(Function(x) x = 99) = False, "Not Exists 99")

        ' RemoveAll
        Dim nums2 As New ArrayList()
        nums2.Add(1)
        nums2.Add(2)
        nums2.Add(3)
        nums2.Add(4)
        nums2.Add(5)
        Dim removed As Integer = nums2.RemoveAll(Function(x) x > 3)
        Assert(removed = 2, "RemoveAll > 3 removed 2")
        Assert(nums2.Count = 3, "RemoveAll leaves 3")

        ' ===== ForEach =====
        Dim total As Integer = 0
        For Each n As Integer In nums
            total = total + n
        Next
        Assert(total = 150, "ForEach sum = 150")

        ' ===== Keyed access (Collection features) =====
        Dim col As New Collection()
        col.Add("Alpha", "a")
        col.Add("Beta", "b")
        col.Add("Gamma", "g")
        Assert(col.Item("a") = "Alpha", "Keyed Item a")
        Assert(col("b") = "Beta", "Keyed indexer b")
        Assert(col.ContainsKey("g") = True, "ContainsKey g")
        col.Remove("b")
        Assert(col.Count = 2, "Remove by key, count = 2")
        Assert(col.ContainsKey("b") = False, "Key b removed")

        ' ===== Summary =====
        Console.WriteLine()
        Console.WriteLine("Passed: " & passed)
        Console.WriteLine("Failed: " & failed)
        Console.WriteLine("Total:  " & (passed + failed))
        If failed = 0 Then
            Console.WriteLine("=== ALL ARRAYLIST TESTS PASSED ===")
        End If
    End Sub
End Module
