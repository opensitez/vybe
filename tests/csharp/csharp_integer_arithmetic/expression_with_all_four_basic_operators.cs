// vybe-test: csharp/csharp_integer_arithmetic/expression_with_all_four_basic_operators
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10 + 6 * 2 - 8 / 4).ToString(), "20");
