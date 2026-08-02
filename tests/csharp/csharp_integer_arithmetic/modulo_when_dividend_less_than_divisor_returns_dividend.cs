// vybe-test: csharp/csharp_integer_arithmetic/modulo_when_dividend_less_than_divisor_returns_dividend
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((4 % 9).ToString(), "4");
