// vybe-test: csharp/csharp_strings_ext/string_builder_replace_substitutes_all_occurrences
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("abab");
sb.Replace("a", "z");
__Check((sb.ToString()).ToString(), "zbzb");
