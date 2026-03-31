' Main VB entry point — uses JS math utilities
Module Program
    Sub Main()
        Console.WriteLine("=== Multi-Language Project ===")
        Console.WriteLine("VB calling shared host functions")

        ' Both VB and JS share the same VM globals and host
        Console.WriteLine("Math.Floor(3.7) = " & CStr(Math.Floor(3.7)))
        Console.WriteLine("Math.Sqrt(16) = " & CStr(Math.Sqrt(16)))

        ' String operations work in both languages
        Dim greeting As String = "Hello from VB!"
        Console.WriteLine(greeting)
        Console.WriteLine("Upper: " & UCase(greeting))
        Console.WriteLine("Length: " & CStr(Len(greeting)))

        Console.WriteLine("=== VB Done ===")
    End Sub
End Module
