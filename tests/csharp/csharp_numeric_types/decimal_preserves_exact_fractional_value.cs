// vybe-test: csharp/csharp_numeric_types/decimal_preserves_exact_fractional_value
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal d = 0.1m + 0.2m; __Check((d).ToString(), "0.3");
