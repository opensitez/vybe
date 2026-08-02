// vybe-test: csharp/csharp_integer_arithmetic/expression_with_addition_subtraction_multiplication
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((2 + 3 * 4 - 5).ToString(), "9");
