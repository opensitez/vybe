// vybe-test: csharp/csharp_integer_arithmetic/division_precedence_over_addition_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10 + 20 / 4).ToString(), "15");
