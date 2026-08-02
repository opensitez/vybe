// vybe-test: csharp/csharp_integer_arithmetic/modulo_with_zero_dividend_yields_zero
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0 % 5).ToString(), "0");
