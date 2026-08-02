// vybe-test: csharp/csharp_integer_arithmetic/expression_with_division_and_modulo_equal_precedence
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((20 / 3 % 2).ToString(), "0");
