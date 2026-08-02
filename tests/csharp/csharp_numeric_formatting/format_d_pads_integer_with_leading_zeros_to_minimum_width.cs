// vybe-test: csharp/csharp_numeric_formatting/format_d_pads_integer_with_leading_zeros_to_minimum_width
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((42.ToString("D5")).ToString(), "00042");
