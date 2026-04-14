Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Module TestSystemFeatures
    Sub Main()
        Console.WriteLine("=== System Features Test Suite ===")
        Console.WriteLine()

        ' ---- System.Random ----
        Console.WriteLine("--- System.Random ---")
        Dim rng As New Random()
        Dim r1 As Integer = rng.Next()
        Console.WriteLine("Random.Next(): " & r1.ToString())
        Dim r2 As Integer = rng.Next(100)
        Console.WriteLine("Random.Next(100): " & r2.ToString())
        Dim r3 As Integer = rng.Next(10, 20)
        Console.WriteLine("Random.Next(10,20): " & r3.ToString())
        Dim rd As Double = rng.NextDouble()
        Console.WriteLine("Random.NextDouble(): " & rd.ToString())

        Dim seeded As New Random(42)
        Dim s1 As Integer = seeded.Next(100)
        Dim s2 As Integer = seeded.Next(100)
        Console.WriteLine("Seeded(42).Next(100) x2: " & s1.ToString() & ", " & s2.ToString())
        Console.WriteLine()

        ' ---- Math extensions ----
        Console.WriteLine("--- Math Extensions ---")
        Console.WriteLine("Math.PI = " & Math.PI.ToString())
        Console.WriteLine("Math.E = " & Math.E.ToString())
        Console.WriteLine("Math.Log10(1000) = " & Math.Log10(1000).ToString())
        Console.WriteLine("Math.Acos(1) = " & Math.Acos(1).ToString())
        Console.WriteLine("Math.Sinh(0) = " & Math.Sinh(0).ToString())
        Console.WriteLine("Math.Cosh(0) = " & Math.Cosh(0).ToString())
        Console.WriteLine("Math.Tanh(0) = " & Math.Tanh(0).ToString())
        Console.WriteLine("Math.Clamp(15, 0, 10) = " & Math.Clamp(15, 0, 10).ToString())
        Console.WriteLine("Math.Clamp(-5, 0, 10) = " & Math.Clamp(-5, 0, 10).ToString())
        Console.WriteLine("Math.Clamp(5, 0, 10) = " & Math.Clamp(5, 0, 10).ToString())
        Console.WriteLine("Math.Log(8, 2) = " & Math.Log(8, 2).ToString())
        Console.WriteLine()

        ' ---- Environment ----
        Console.WriteLine("--- Environment ---")
        Dim cd As String = Environment.CurrentDirectory
        Console.WriteLine("CurrentDirectory: " & cd)
        Dim mn As String = Environment.MachineName
        Console.WriteLine("MachineName: " & mn)
        Dim un As String = Environment.UserName
        Console.WriteLine("UserName: " & un)
        Dim osv As String = Environment.OSVersion
        Console.WriteLine("OSVersion: " & osv)
        Dim pc As Integer = Environment.ProcessorCount
        Console.WriteLine("ProcessorCount: " & pc.ToString())
        Dim is64 As Boolean = Environment.Is64BitOperatingSystem
        Console.WriteLine("Is64BitOS: " & is64.ToString())
        Dim nl As String = Environment.NewLine
        Console.WriteLine("NewLine length: " & nl.Length.ToString())

        Environment.SetEnvironmentVariable("vybe_TEST_VAR", "hello_vybe")
        Dim tv As String = Environment.GetEnvironmentVariable("vybe_TEST_VAR")
        Console.WriteLine("GetEnvVar(vybe_TEST_VAR): " & tv)

        Dim desktop As String = Environment.GetFolderPath(0)
        Console.WriteLine("GetFolderPath(Desktop): " & desktop)
        Console.WriteLine()

        ' ---- Guid ----
        Console.WriteLine("--- Guid ---")
        Dim g1 As String = Guid.NewGuid().ToString()
        Console.WriteLine("Guid.NewGuid: " & g1)
        Dim g2 As String = Guid.NewGuid().ToString()
        Console.WriteLine("Guid.NewGuid (different): " & g2)
        Dim empty As String = Guid.Empty
        Console.WriteLine("Guid.Empty: " & empty)
        Console.WriteLine("Guids differ: " & (g1 <> g2).ToString())
        Console.WriteLine()

        ' ---- Convert Base64 ----
        Console.WriteLine("--- Convert Base64 ---")
        Dim original As String = "Hello, World!"
        Dim encoded As String = Convert.ToBase64String(original)
        Console.WriteLine("ToBase64: " & encoded)
        Dim decoded As String = Convert.FromBase64String(encoded)
        Console.WriteLine("FromBase64: " & decoded)
        Console.WriteLine("Roundtrip OK: " & (original = decoded).ToString())

        Dim encoded2 As String = Convert.ToBase64String("ABCDEF")
        Console.WriteLine("ToBase64(ABCDEF): " & encoded2)
        Dim decoded2 As String = Convert.FromBase64String(encoded2)
        Console.WriteLine("FromBase64 roundtrip: " & decoded2)
        Console.WriteLine()

        ' ---- Stopwatch ----
        Console.WriteLine("--- Stopwatch ---")
        Dim sw As New Stopwatch()
        sw.Start()
        ' Do some work
        Dim dummy As Integer = 0
        For i As Integer = 1 To 10000
            dummy = dummy + i
        Next
        sw.Stop()
        Dim elapsed As Long = sw.ElapsedMilliseconds
        Console.WriteLine("Stopwatch elapsed (ms): " & elapsed.ToString())
        Console.WriteLine("Stopwatch >= 0: " & (elapsed >= 0).ToString())

        sw.Reset()
        Console.WriteLine("After Reset: " & sw.ElapsedMilliseconds.ToString())

        Dim sw2 As Stopwatch = Stopwatch.StartNew()
        For i As Integer = 1 To 5000
            dummy = dummy + i
        Next
        sw2.Stop()
        Console.WriteLine("StartNew elapsed >= 0: " & (sw2.ElapsedMilliseconds >= 0).ToString())
        Console.WriteLine()

        ' ---- Path extensions ----
        Console.WriteLine("--- Path Extensions ---")
        Console.WriteLine("GetFileNameWithoutExtension: " & Path.GetFileNameWithoutExtension("/foo/bar/test.txt"))
        Console.WriteLine("GetTempPath: " & Path.GetTempPath())
        Console.WriteLine("HasExtension(.txt): " & Path.HasExtension("file.txt").ToString())
        Console.WriteLine("HasExtension(noext): " & Path.HasExtension("noext").ToString())
        Console.WriteLine("IsPathRooted(/abs): " & Path.IsPathRooted("/abs/path").ToString())
        Console.WriteLine("IsPathRooted(rel): " & Path.IsPathRooted("rel/path").ToString())
        Console.WriteLine("ChangeExtension: " & Path.ChangeExtension("test.txt", ".md"))
        Console.WriteLine("GetExtension: " & Path.GetExtension("hello.vb"))
        Console.WriteLine("GetFileName: " & Path.GetFileName("/a/b/c.txt"))
        Console.WriteLine("GetDirectoryName: " & Path.GetDirectoryName("/a/b/c.txt"))
        Console.WriteLine()

        ' ---- StreamReader / StreamWriter ----
        Console.WriteLine("--- StreamReader / StreamWriter ---")
        Dim testFile As String = "/tmp/vybe_stream_test.txt"
        Dim writer As New StreamWriter(testFile)
        writer.WriteLine("Line one")
        writer.WriteLine("Line two")
        writer.Write("Line three")
        writer.Close()

        Dim reader As New StreamReader(testFile)
        Dim line1 As String = reader.ReadLine()
        Console.WriteLine("Read line 1: " & line1)
        Dim line2 As String = reader.ReadLine()
        Console.WriteLine("Read line 2: " & line2)
        Dim rest As String = reader.ReadToEnd()
        Console.WriteLine("ReadToEnd: " & rest)
        Console.WriteLine("EndOfStream: " & reader.EndOfStream.ToString())
        reader.Close()
        Console.WriteLine()

        ' ---- Regex instance ----
        Console.WriteLine("--- Regex Instance ---")
        Dim rx As New Regex("\d+")
        Console.WriteLine("IsMatch(abc123): " & rx.IsMatch("abc123").ToString())
        Console.WriteLine("IsMatch(abcdef): " & rx.IsMatch("abcdef").ToString())

        Dim m As String = rx.Match("hello42world")
        Console.WriteLine("Match(hello42world): " & m)

        Dim replaced As String = rx.Replace("a1b2c3", "X")
        Console.WriteLine("Replace digits with X: " & replaced)

        Dim rxI As New Regex("hello", RegexOptions.IgnoreCase)
        Console.WriteLine("IgnoreCase IsMatch(HELLO): " & rxI.IsMatch("HELLO WORLD").ToString())
        Console.WriteLine()

        ' ---- List methods ----
        Console.WriteLine("--- List Methods ---")
        Dim nums As New List(Of Integer)
        nums.Add(5)
        nums.Add(3)
        nums.Add(8)
        nums.Add(1)
        nums.Add(3)
        Console.WriteLine("Before sort: " & nums(0).ToString() & "," & nums(1).ToString() & "," & nums(2).ToString() & "," & nums(3).ToString() & "," & nums(4).ToString())

        nums.Sort()
        Console.WriteLine("After sort: " & nums(0).ToString() & "," & nums(1).ToString() & "," & nums(2).ToString() & "," & nums(3).ToString() & "," & nums(4).ToString())

        nums.Reverse()
        Console.WriteLine("After reverse: " & nums(0).ToString() & "," & nums(1).ToString() & "," & nums(2).ToString() & "," & nums(3).ToString() & "," & nums(4).ToString())

        Console.WriteLine("IndexOf(3): " & nums.IndexOf(3).ToString())
        Console.WriteLine("LastIndexOf(3): " & nums.LastIndexOf(3).ToString())
        Console.WriteLine("Contains(8): " & nums.Contains(8).ToString())
        Console.WriteLine("Contains(99): " & nums.Contains(99).ToString())

        nums.Insert(2, 42)
        Console.WriteLine("After Insert(2, 42): " & nums(2).ToString())

        Dim more As New List(Of Integer)
        more.Add(100)
        more.Add(200)
        nums.AddRange(more)
        Console.WriteLine("After AddRange, Count: " & nums.Count.ToString())
        Console.WriteLine()

        ' ---- Dictionary.TryGetValue ----
        Console.WriteLine("--- Dictionary.TryGetValue ---")
        Dim dict As New Dictionary(Of String, Integer)
        dict.Add("alpha", 1)
        dict.Add("beta", 2)
        Dim outVal As Integer = 0
        Dim found As Boolean = dict.TryGetValue("alpha", outVal)
        Console.WriteLine("TryGetValue(alpha): " & found.ToString() & " = " & outVal.ToString())
        Dim found2 As Boolean = dict.TryGetValue("gamma", outVal)
        Console.WriteLine("TryGetValue(gamma): " & found2.ToString() & " = " & outVal.ToString())
        Console.WriteLine()

        ' ---- Environment.TickCount ----
        Console.WriteLine("--- TickCount ---")
        Dim tc As Long = Environment.TickCount
        Console.WriteLine("TickCount > 0: " & (tc > 0).ToString())
        Console.WriteLine()

        Console.WriteLine("=== All Tests Passed ===")
    End Sub
End Module
