// vybe-test: csharp/csharp_string_interpolation/format_specifier_in_interpolation_applies_number_format
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d=1234.5; __Check(($"{d:N2}").ToString(), "1,234.50");
