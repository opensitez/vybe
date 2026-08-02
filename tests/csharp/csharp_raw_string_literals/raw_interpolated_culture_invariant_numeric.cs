// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_culture_invariant_numeric
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double value=1234.5; string text=$"""{value.ToString(System.Globalization.CultureInfo.InvariantCulture)}"""; __Check((text.Contains(".")).ToString(), "True");
