// vybe-test: csharp/csharp_numeric_formatting/format_g_chooses_shorter_of_fixed_or_scientific_representation
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((0.00001).ToString("G")).ToString(), "1E-05");
