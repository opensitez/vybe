// vybe-test: csharp/csharp_integer_arithmetic/subtraction_after_addition_in_single_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((2 + 8 - 5).ToString(), "5");
