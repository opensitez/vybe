// vybe-test: csharp/csharp_numeric_checked_bitwise/modulo_operator_returns_remainder_after_division
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((29 % 6).ToString(), "5");
