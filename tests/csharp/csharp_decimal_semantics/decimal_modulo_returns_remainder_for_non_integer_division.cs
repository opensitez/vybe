// vybe-test: csharp/csharp_decimal_semantics/decimal_modulo_returns_remainder_for_non_integer_division
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10.5m % 3m).ToString(), "1.5");
