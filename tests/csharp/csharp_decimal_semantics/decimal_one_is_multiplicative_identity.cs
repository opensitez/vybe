// vybe-test: csharp/csharp_decimal_semantics/decimal_one_is_multiplicative_identity
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal value = 9.75m; __Check((value * 1m).ToString(), "9.75");
