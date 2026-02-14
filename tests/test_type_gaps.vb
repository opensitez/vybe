Imports System
Imports System.IO
Imports System.Collections.Generic

Module TestTypeGaps
    Sub Main()
        TestMath()
        TestFile()
        TestDirectory()
        TestPath()
        TestString()
        TestDictionary()
        TestDateTime()
    End Sub

    Sub TestMath()
        Console.WriteLine("--- TestMath ---")
        Console.WriteLine("Math.Abs(-5): " & Math.Abs(-5))
        Console.WriteLine("Math.Round(3.14159, 2): " & Math.Round(3.14159, 2))
        Console.WriteLine("Math.Min(10, 20): " & Math.Min(10, 20))
        Console.WriteLine("Math.Max(10, 20): " & Math.Max(10, 20))
        Console.WriteLine("Math.Sign(-10): " & Math.Sign(-10))
        Console.WriteLine("Math.Sqrt(16): " & Math.Sqrt(16))
    End Sub

    Sub TestFile()
        Console.WriteLine("--- TestFile ---")
        Dim f = "test_file_gaps.txt"
        If File.Exists(f) Then File.Delete(f)
        
        File.WriteAllText(f, "Hello File Gaps")
        Console.WriteLine("File.Exists: " & File.Exists(f))
        
        Dim content = File.ReadAllText(f)
        Console.WriteLine("Content: " & content)
        
        File.AppendAllText(f, " Appended")
        Console.WriteLine("Appended Content: " & File.ReadAllText(f))
        
        File.Delete(f)
        Console.WriteLine("File.Exists after delete: " & File.Exists(f))
    End Sub

    Sub TestDirectory()
        Console.WriteLine("--- TestDirectory ---")
        Dim d = "test_dir_gaps"
        If Directory.Exists(d) Then Directory.Delete(d)
        
        Directory.CreateDirectory(d)
        Console.WriteLine("Directory.Exists: " & Directory.Exists(d))
        
        Dim files = Directory.GetFiles(".")
        Console.WriteLine("GetFiles count > 0: " & (files.Length > 0))
        
        Directory.Delete(d)
        Console.WriteLine("Directory.Exists after delete: " & Directory.Exists(d))
    End Sub

    Sub TestPath()
        Console.WriteLine("--- TestPath ---")
        Dim p = Path.Combine("folder", "file.txt")
        Console.WriteLine("Combine: " & p)
        Console.WriteLine("GetFileName: " & Path.GetFileName(p))
        Console.WriteLine("GetExtension: " & Path.GetExtension(p))
        Console.WriteLine("ChangeExtension: " & Path.ChangeExtension(p, "md"))
    End Sub

    Sub TestString()
        Console.WriteLine("--- TestString ---")
        Dim s = "Hello"
        Console.WriteLine("PadLeft(10, '-'): '" & s.PadLeft(10, "-") & "'")
        Console.WriteLine("PadRight(10, '-'): '" & s.PadRight(10, "-") & "'")
        
        Dim sentence = "Hello World Vybe"
        Dim parts = sentence.Split(" ")
        Console.WriteLine("Split count: " & parts.Length)
        Console.WriteLine("Part 0: " & parts(0))
        Console.WriteLine("Part 2: " & parts(2))
    End Sub

    Sub TestDictionary()
        Console.WriteLine("--- TestDictionary ---")
        Dim d As New Dictionary(Of String, Integer)
        d.Add("One", 1)
        d.Add("Two", 2)
        
        Console.WriteLine("Keys count: " & d.Keys.Length)
        Console.WriteLine("Values count: " & d.Values.Length)
        Console.WriteLine("ContainsKey One: " & d.ContainsKey("One"))
        
        d.Remove("One")
        Console.WriteLine("ContainsKey One after remove: " & d.ContainsKey("One"))
        Console.WriteLine("Count: " & d.Count)
    End Sub
    
    Sub TestDateTime()
        Console.WriteLine("--- TestDateTime ---")
        ' Create a specific date 2023-10-05 14:30:00
        ' We don't have DateTime constructor exposed via New DateTime yet?
        ' interpreter.rs line 3300 handles arguments for New? 
        ' Actually DateTime is a struct, often New DateTime(y,m,d) works if mapped.
        ' If not, we use DateSerial equivalent or just Now.
        
        Dim d = DateTime.Now
        ' Test standard formats
        Console.WriteLine("d format: " & d.ToString("d"))
        Console.WriteLine("yyyy-MM-dd format: " & d.ToString("yyyy-MM-dd"))
        
        Dim d2 = d.AddDays(1)
        Console.WriteLine("AddDays(1) > Now: " & (d2 > d))
        
        Dim diff = d2.Subtract(d)
        Console.WriteLine("Subtract TotalDays: " & diff.TotalDays)
    End Sub
End Module
