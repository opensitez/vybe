// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_with_format_specifier_pads_numeric_output
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 7; __Check(($"{n:D3}").ToString(), "007");
