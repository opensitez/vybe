// vybe-test: csharp/csharp_decimal_semantics/decimal_to_string_preserves_trailing_zero_from_format
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal value = 3.5m; __Check((value.ToString("0.00")).ToString(), "3.50");
