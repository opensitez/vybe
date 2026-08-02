// vybe-test: csharp/csharp_string_culture/string_compare_invariant_culture_ignores_locale
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int r=string.Compare("hello","HELLO",System.StringComparison.InvariantCultureIgnoreCase);
__Check((r==0).ToString(), "True");
