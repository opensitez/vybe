// vybe-test: csharp/csharp_parse_with_invariant_culture/decimal_to_string_invariant_uses_dot_not_comma_for_fractions
// origin: languages/csharp/tests/csharp/test_csharp_parse_with_invariant_culture.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal value = 2.25m;
__Check((value.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString(), "2.25");
