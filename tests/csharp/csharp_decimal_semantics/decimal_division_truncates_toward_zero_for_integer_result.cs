// vybe-test: csharp/csharp_decimal_semantics/decimal_division_truncates_toward_zero_for_integer_result
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal total = 10m; decimal parts = 4m; __Check((total / parts).ToString(), "2.5");
