// vybe-test: csharp/csharp_numeric_formatting/format_e_expresses_double_in_scientific_notation
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((1000.0).ToString("E2")).ToString(), "1.00E+003");
