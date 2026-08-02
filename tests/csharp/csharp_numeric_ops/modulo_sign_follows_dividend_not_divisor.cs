// vybe-test: csharp/csharp_numeric_ops/modulo_sign_follows_dividend_not_divisor
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((7%3).ToString(), "1");
__Check((-7%3).ToString(), "-1");
