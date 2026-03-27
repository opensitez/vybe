Module Program
    Sub Main()
        Console.WriteLine("=== VB calling Rust WASM ===")
        Console.WriteLine("add(3, 4) = " & CStr(add(3, 4)))
        Console.WriteLine("add(100, 200) = " & CStr(add(100, 200)))
        Console.WriteLine("add(-5, 15) = " & CStr(add(-5, 15)))
        Console.WriteLine("=== Done ===")
    End Sub
End Module
