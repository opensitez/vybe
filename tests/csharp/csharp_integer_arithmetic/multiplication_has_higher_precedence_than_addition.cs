// vybe-test: csharp/csharp_integer_arithmetic/multiplication_has_higher_precedence_than_addition
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((2 + 3 * 4).ToString(), "14");
