// vybe-test: csharp/csharp_numeric_formatting/format_custom_zero_placeholder_pads_fractional_digits
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((1.5).ToString("0.00")).ToString(), "1.50");
