// vybe-test: csharp/csharp_numeric_checked_bitwise/decimal_addition_preserves_decimal_precision
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal left = 1.2m; decimal right = 2.3m; __Check((left + right).ToString(), "3.5");
