// vybe-test: csharp/csharp_string_ops_advanced/interpolated_string_with_format_specifier
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double pi=3.14159; __Check(($"{pi:F2}").ToString(), "3.14");
