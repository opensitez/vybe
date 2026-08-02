// vybe-test: csharp/csharp_comparison_sorting/dictionary_with_case_insensitive_comparer_finds_key_variant
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase) { ["Key"] = 3 }; __Check((map.ContainsKey("key")).ToString(), "True");
