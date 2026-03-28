// Fibonacci in C# — demonstrates recursion, loops, and classes
// Run: vybec examples/csharp/fibonacci.cs

using System;

class Fibonacci
{
    // Recursive
    public static int Recursive(int n)
    {
        if (n <= 1) return n;
        return Recursive(n - 1) + Recursive(n - 2);
    }

    // Iterative
    public static int Iterative(int n)
    {
        if (n <= 1) return n;
        int a = 0;
        int b = 1;
        for (int i = 2; i <= n; i++)
        {
            int temp = a + b;
            a = b;
            b = temp;
        }
        return b;
    }

    static void Main()
    {
        Console.WriteLine("Fibonacci (recursive):");
        for (int i = 0; i <= 10; i++)
        {
            Console.WriteLine($"  F({i}) = {Recursive(i)}");
        }

        Console.WriteLine("Fibonacci (iterative):");
        for (int i = 0; i <= 20; i++)
        {
            Console.WriteLine($"  F({i}) = {Iterative(i)}");
        }
    }
}
