Module StringBuilderTest
    Sub Main()
        Console.WriteLine("=== StringBuilder Test Start ===")

        ' --- Basic construction and Append ---
        Dim sb As New StringBuilder()
        sb.Append("Hello")
        sb.Append(" ")
        sb.Append("World")
        Console.WriteLine("Append: " & sb.ToString())

        ' --- Constructor with initial string ---
        Dim sb2 As New StringBuilder("Init")
        Console.WriteLine("InitCtor: " & sb2.ToString())

        ' --- Constructor with capacity ---
        Dim sb3 As New StringBuilder(64)
        sb3.Append("cap")
        Console.WriteLine("CapCtor: " & sb3.ToString())

        ' --- AppendLine ---
        Dim sbLine As New StringBuilder()
        sbLine.Append("Line1")
        sbLine.AppendLine()
        sbLine.Append("Line2")
        Dim lineResult = sbLine.ToString()
        ' Check it contains both lines
        Console.WriteLine("AppendLine Contains Line1: " & lineResult.Contains("Line1"))
        Console.WriteLine("AppendLine Contains Line2: " & lineResult.Contains("Line2"))

        ' --- AppendFormat ---
        Dim sbFmt As New StringBuilder()
        sbFmt.AppendFormat("Name: {0}, Age: {1}", "Alice", 30)
        Console.WriteLine("AppendFormat: " & sbFmt.ToString())

        ' --- Insert ---
        Dim sbIns As New StringBuilder("HelloWorld")
        sbIns.Insert(5, " ")
        Console.WriteLine("Insert: " & sbIns.ToString())

        ' --- Remove ---
        Dim sbRem As New StringBuilder("Hello World")
        sbRem.Remove(5, 6)
        Console.WriteLine("Remove: " & sbRem.ToString())

        ' --- Replace ---
        Dim sbRep As New StringBuilder("Hello World")
        sbRep.Replace("World", "VB")
        Console.WriteLine("Replace: " & sbRep.ToString())

        ' --- Clear ---
        Dim sbClr As New StringBuilder("SomeText")
        sbClr.Clear()
        Console.WriteLine("Clear Length: " & sbClr.Length)
        Console.WriteLine("Clear ToString: '" & sbClr.ToString() & "'")

        ' --- Length property read ---
        Dim sbLen As New StringBuilder("ABCDE")
        Console.WriteLine("Length: " & sbLen.Length)

        ' --- Length property set (truncate) ---
        Dim sbTrunc As New StringBuilder("Hello World")
        sbTrunc.Length = 5
        Console.WriteLine("Truncate: " & sbTrunc.ToString())

        ' --- Length property set (pad with nulls then overwrite) ---
        Dim sbPad As New StringBuilder("Hi")
        sbPad.Length = 5
        Console.WriteLine("PadLength: " & sbPad.Length)

        ' --- Method chaining ---
        Dim sbChain As New StringBuilder()
        sbChain.Append("A").Append("B").Append("C")
        Console.WriteLine("Chain: " & sbChain.ToString())

        ' --- EnsureCapacity ---
        Dim sbCap As New StringBuilder()
        sbCap.EnsureCapacity(100)
        Dim cap = sbCap.Capacity
        Console.WriteLine("EnsureCapacity >= 100: " & (cap >= 100))

        ' --- Chars (indexer) ---
        Dim sbChar As New StringBuilder("ABCDE")
        Dim ch = sbChar.Chars(2)
        Console.WriteLine("Chars(2): " & ch)

        ' --- ToString ---
        Dim sbTs As New StringBuilder("Final")
        Console.WriteLine("ToString: " & sbTs.ToString())

        ' --- Equals ---
        Dim sbEq1 As New StringBuilder("Same")
        Dim sbEq2 As New StringBuilder("Same")
        Dim sbEq3 As New StringBuilder("Diff")
        Console.WriteLine("Equals Same: " & sbEq1.Equals(sbEq2))
        Console.WriteLine("Equals Diff: " & sbEq1.Equals(sbEq3))

        ' --- CopyTo ---
        Dim sbCopy As New StringBuilder("Hello World")
        Dim dest(10) As Char
        sbCopy.CopyTo(0, dest, 0, 5)
        ' dest should have H, e, l, l, o in first 5 positions
        Console.WriteLine("CopyTo(0): " & dest(0))
        Console.WriteLine("CopyTo(4): " & dest(4))

        ' --- Complex scenario: build a CSV line ---
        Dim sbCsv As New StringBuilder()
        sbCsv.Append("Name").Append(",").Append("Age").Append(",").Append("City")
        Console.WriteLine("CSV: " & sbCsv.ToString())

        ' --- Replace chained ---
        Dim sbRepChain As New StringBuilder("aXbXc")
        sbRepChain.Replace("X", "-")
        Console.WriteLine("ReplaceChain: " & sbRepChain.ToString())

        ' --- Full qualified name ---
        Dim sbFull As New System.Text.StringBuilder()
        sbFull.Append("FullQualified")
        Console.WriteLine("FullQualified: " & sbFull.ToString())

        Console.WriteLine("=== StringBuilder Test End ===")
    End Sub
End Module
