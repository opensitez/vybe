// vybe-test: csharp/csharp_integer_arithmetic/deeply_nested_parentheses_resolve_innermost_first
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((((2 + 3) * (4 - 1)) / 3).ToString(), "5");
