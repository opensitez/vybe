// vybe-test: csharp/csharp_numeric_precision/decimal_preserves_trailing_zeros_in_precision
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal d=1.50m;
__Check((d.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString(), "1.50");
