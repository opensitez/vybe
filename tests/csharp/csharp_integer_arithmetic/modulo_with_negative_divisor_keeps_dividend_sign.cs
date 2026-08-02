// vybe-test: csharp/csharp_integer_arithmetic/modulo_with_negative_divisor_keeps_dividend_sign
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10 % -3).ToString(), "1");
