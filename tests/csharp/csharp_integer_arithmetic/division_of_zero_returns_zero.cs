// vybe-test: csharp/csharp_integer_arithmetic/division_of_zero_returns_zero
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0 / 9).ToString(), "0");
