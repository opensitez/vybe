' Call Rust WASM functions from VB
Module Program
    Sub Main()
        Console.WriteLine("=== VB calling Rust WASM ===")

        Console.WriteLine("add(3, 4) = " & CStr(add(3, 4)))
        Console.WriteLine("add(100, 200) = " & CStr(add(100, 200)))

        Console.WriteLine("factorial(5) = " & CStr(factorial(5)))
        Console.WriteLine("factorial(10) = " & CStr(factorial(10)))

        Console.WriteLine("fibonacci(10) = " & CStr(fibonacci(10)))
        Console.WriteLine("fibonacci(20) = " & CStr(fibonacci(20)))

        Console.WriteLine("is_prime(17) = " & CStr(is_prime(17)))
        Console.WriteLine("is_prime(18) = " & CStr(is_prime(18)))
        Console.WriteLine("is_prime(97) = " & CStr(is_prime(97)))

        Console.WriteLine("=== Done ===")
    End Sub
End Module
