// vybe-test: csharp/csharp_numeric_formatting/format_f2_rounds_double_to_two_decimal_places
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((3.14159).ToString("F2")).ToString(), "3.14");
