// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_with_format_specifier
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double pi=3.14159; string text=$"""pi={pi:F2}"""; __Check((text).ToString(), "pi=3.14");
