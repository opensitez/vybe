// vybe-test: csharp/csharp_decimal_semantics/decimal_addition_preserves_fractional_sum_without_binary_drift
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal a = 0.1m; decimal b = 0.2m; __Check((a + b).ToString(), "0.3");
