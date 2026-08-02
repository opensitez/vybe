// vybe-test: csharp/csharp_integer_arithmetic/negative_literal_in_multiplication_and_addition
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((-3 * 4 + 10).ToString(), "-2");
