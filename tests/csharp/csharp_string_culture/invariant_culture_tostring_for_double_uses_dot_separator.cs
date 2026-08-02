// vybe-test: csharp/csharp_string_culture/invariant_culture_tostring_for_double_uses_dot_separator
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d=1.5;
__Check((d.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString(), "1.5");
