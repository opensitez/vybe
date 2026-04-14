' Test System.IO.File
Dim path1 As String
path1 = TestFilePath
Console.WriteLine("Path1: '" & path1 & "'")
System.IO.File.WriteAllText(path1, "Hello from .NET I/O")

Dim content1 As String
content1 = System.IO.File.ReadAllText(path1)
Console.WriteLine("NetRead: " & content1)

If System.IO.File.Exists(path1) Then
    Console.WriteLine("NetExists: True")
End If

' Test System.IO.Path
Dim dir As String = System.IO.Path.GetDirectoryName("/foo/bar/baz.txt")
Console.WriteLine("PathDir: " & dir)
Dim ext As String = System.IO.Path.GetExtension("file.txt")
Console.WriteLine("PathExt: " & ext)

' Test Legacy I/O
Dim path2 As String = TestLegacyPath
Dim f As Integer = 1

Open path2 For Output As #f
Print #f, "Line 1"
Print #f, "Line 2"
Close #f

Open path2 For Input As #f
Dim line1 As String
Dim line2 As String
Line Input #f, line1
Line Input #f, line2
Close #f

Console.WriteLine("LegacyRead1: " & line1)
Console.WriteLine("LegacyRead2: " & line2)

' Cleanup
System.IO.File.Delete(path1)
System.IO.File.Delete(path2)
