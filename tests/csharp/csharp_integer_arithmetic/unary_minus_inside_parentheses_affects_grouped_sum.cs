// vybe-test: csharp/csharp_integer_arithmetic/unary_minus_inside_parentheses_affects_grouped_sum
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((-(3 + 4)).ToString(), "-7");
