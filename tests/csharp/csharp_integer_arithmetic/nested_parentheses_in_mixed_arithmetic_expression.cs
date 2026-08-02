// vybe-test: csharp/csharp_integer_arithmetic/nested_parentheses_in_mixed_arithmetic_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((8 - 2) * (3 + 1)).ToString(), "24");
