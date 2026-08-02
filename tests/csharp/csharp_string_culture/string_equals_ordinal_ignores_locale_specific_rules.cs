// vybe-test: csharp/csharp_string_culture/string_equals_ordinal_ignores_locale_specific_rules
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("Abc".Equals("abc",System.StringComparison.OrdinalIgnoreCase)).ToString(), "True");
